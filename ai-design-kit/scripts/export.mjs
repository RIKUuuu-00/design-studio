#!/usr/bin/env node
// アートボード（designs/<slug>/*.html）を画面幅ごとに PNG（と任意で PDF）へ書き出す。
//
//   node export.mjs designs/settings [--widths 375,1280] [--only a,b] [--pdf] [--out shots]
//
// - index.html（比較キャンバス）は対象外
// - Playwright はプロジェクト / グローバル / PLAYWRIGHT_MODULE の順に探す
// - 社内プロキシで Chromium を取得できない場合は PLAYWRIGHT_CHANNEL=chrome か msedge で既存ブラウザを使う
//   （PLAYWRIGHT_EXECUTABLE_PATH でブラウザ本体を直接指定してもよい）

import { createRequire } from "node:module";
import { execSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

function parseArgs(argv) {
  const args = { dir: null, widths: [375, 1280], only: null, pdf: false, out: "shots", height: 900 };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--widths") args.widths = argv[++i].split(",").map(Number);
    else if (a === "--only") args.only = argv[++i].split(",");
    else if (a === "--pdf") args.pdf = true;
    else if (a === "--out") args.out = argv[++i];
    else if (a === "--height") args.height = Number(argv[++i]);
    else if (!args.dir) args.dir = a;
  }
  if (!args.dir) {
    console.error("使い方: node export.mjs designs/<slug> [--widths 375,1280] [--only a,b] [--pdf]");
    process.exit(2);
  }
  return args;
}

function loadPlaywright() {
  const candidates = [process.env.PLAYWRIGHT_MODULE, path.join(process.cwd(), "package.json")];
  try { candidates.push(path.join(execSync("npm root -g", { encoding: "utf8" }).trim(), "noop.js")); } catch {}
  for (const c of candidates.filter(Boolean)) {
    try {
      const req = createRequire(c.endsWith(".js") || c.endsWith(".json") ? c : path.join(c, "noop.js"));
      return req("playwright");
    } catch {}
  }
  console.error("playwright が見つからない。`npm i -D playwright` か `npm i -g playwright` を実行すること。");
  process.exit(2);
}

const args = parseArgs(process.argv.slice(2));
const dir = path.resolve(args.dir);
const files = fs.readdirSync(dir)
  .filter((f) => f.endsWith(".html") && f !== "index.html")
  .filter((f) => !args.only || args.only.includes(path.basename(f, ".html")))
  .sort();
if (!files.length) {
  console.error(`アートボードが無い: ${dir}`);
  process.exit(1);
}
const outDir = path.resolve(dir, args.out);
fs.mkdirSync(outDir, { recursive: true });

const { chromium } = loadPlaywright();
const launch = {};
if (process.env.PLAYWRIGHT_CHANNEL) launch.channel = process.env.PLAYWRIGHT_CHANNEL;
if (process.env.PLAYWRIGHT_EXECUTABLE_PATH) launch.executablePath = process.env.PLAYWRIGHT_EXECUTABLE_PATH;
const browser = await chromium.launch(launch);
try {
  for (const f of files) {
    const id = path.basename(f, ".html");
    const url = pathToFileURL(path.join(dir, f)).href;
    for (const w of args.widths) {
      const page = await browser.newPage({ viewport: { width: w, height: args.height }, deviceScaleFactor: 1 });
      const errors = [];
      page.on("pageerror", (e) => errors.push(e.message));
      page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
      await page.goto(url, { waitUntil: "networkidle" });
      await page.evaluate(() => document.fonts && document.fonts.ready);
      // 横スクロールの発生（レイアウト崩れの代表例）を検出する
      const overflowX = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
      const target = path.join(outDir, `${id}-${w}.png`);
      await page.screenshot({ path: target, fullPage: true });
      const notes = [];
      if (overflowX > 1) notes.push(`横はみ出し ${overflowX}px`);
      if (errors.length) notes.push(`コンソールエラー ${errors.length}件: ${errors[0]}`);
      console.log(`png: ${path.relative(process.cwd(), target)}${notes.length ? "  [" + notes.join(" / ") + "]" : ""}`);
      await page.close();
    }
    if (args.pdf) {
      const page = await browser.newPage({ viewport: { width: Math.max(...args.widths), height: args.height } });
      await page.goto(url, { waitUntil: "networkidle" });
      const target = path.join(outDir, `${id}.pdf`);
      await page.pdf({ path: target, printBackground: true, width: `${Math.max(...args.widths)}px` });
      console.log(`pdf: ${path.relative(process.cwd(), target)}`);
      await page.close();
    }
  }
} finally {
  await browser.close();
}
