#!/usr/bin/env node
// デザイントークン外の値（色・余白・角丸・文字サイズ・フォント）を検出する。
//
//   node token-lint.mjs designs/settings [--tokens design-system/tokens.css] [--strict]
//
// - 対象: 指定ディレクトリ配下の .html / .css（tokens.css 自身と index.html は除外）
// - 色: #hex / rgb() / hsl() の直書き → var(--color-*) を使う
// - 長さ: margin / padding / gap / border-radius / font-size / top 等の px・rem 直書き → var(--space-*) 等
//   （0、1px の罫線、100% などの相対値は許容）
// - font-family の直書き → var(--font-*)
// - 未定義のトークン参照 var(--xxx)（タイプミス）
// 既定は終了コード 0。--strict で指摘があれば 1。

import fs from "node:fs";
import path from "node:path";

const argv = process.argv.slice(2);
let target = null;
let tokensPath = "design-system/tokens.css";
let strict = false;
for (let i = 0; i < argv.length; i++) {
  if (argv[i] === "--tokens") tokensPath = argv[++i];
  else if (argv[i] === "--strict") strict = true;
  else target = argv[i];
}
if (!target) {
  console.error("使い方: node token-lint.mjs <dir|file> [--tokens design-system/tokens.css] [--strict]");
  process.exit(2);
}
if (!fs.existsSync(tokensPath)) {
  console.error(`tokens.css が無い: ${tokensPath}`);
  process.exit(2);
}

const tokensCss = fs.readFileSync(tokensPath, "utf8");
const defined = new Set([...tokensCss.matchAll(/(--[\w-]+)\s*:/g)].map((m) => m[1]));

function walk(p) {
  const st = fs.statSync(p);
  if (st.isFile()) return [p];
  return fs.readdirSync(p).flatMap((f) => (f === "node_modules" || f.startsWith(".") ? [] : walk(path.join(p, f))));
}

const files = walk(target).filter((f) => /\.(html|css)$/.test(f))
  .filter((f) => path.resolve(f) !== path.resolve(tokensPath) && path.basename(f) !== "index.html");

const LENGTH_PROPS = /^(margin|padding|gap|row-gap|column-gap|border-radius|font-size|line-height|top|right|bottom|left|inset|width|height|min-width|max-width|min-height|max-height)(-[\w-]+)?$/;
const SIZE_PROPS_STRICT = /^(margin|padding|gap|row-gap|column-gap|border-radius|font-size)(-[\w-]+)?$/;
const COLOR_RE = /#[0-9a-fA-F]{3,8}\b|\b(?:rgba?|hsla?)\(/;
const findings = [];

function extractCss(file, text) {
  if (file.endsWith(".css")) return [{ css: text, offset: 0 }];
  const blocks = [];
  for (const m of text.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/g)) blocks.push({ css: m[1], offset: m.index + m[0].indexOf(m[1]) });
  for (const m of text.matchAll(/\sstyle="([^"]*)"/g)) blocks.push({ css: m[1], offset: m.index + m[0].indexOf(m[1]), inline: true });
  return blocks;
}

function lineOf(text, index) {
  return text.slice(0, index).split("\n").length;
}

for (const file of files) {
  const text = fs.readFileSync(file, "utf8");
  for (const block of extractCss(file, text)) {
    const css = block.css.replace(/\/\*[\s\S]*?\*\//g, (c) => " ".repeat(c.length));
    for (const m of css.matchAll(/([a-zA-Z-]+)\s*:\s*([^;{}]+)/g)) {
      const prop = m[1].toLowerCase();
      const value = m[2].trim();
      if (prop.startsWith("--")) continue; // ローカル変数の定義は対象外
      const line = lineOf(text, block.offset + m.index);
      const at = `${path.relative(process.cwd(), file)}:${line}`;
      const noVars = value.replace(/var\([^)]*\)/g, "");
      if (COLOR_RE.test(noVars)) findings.push({ at, rule: "color", msg: `${prop}: ${value} → 色は var(--color-*) を使う` });
      if (prop === "font-family" && !/var\(/.test(value) && !/^(inherit|initial|unset)$/.test(value)) {
        findings.push({ at, rule: "font", msg: `font-family: ${value} → var(--font-*) を使う` });
      }
      if (SIZE_PROPS_STRICT.test(prop) || (LENGTH_PROPS.test(prop) && block.inline)) {
        const raw = [...noVars.matchAll(/(-?\d*\.?\d+)(px|rem|em)\b/g)].filter(([, n, u]) => !(Number(n) === 0 || (u === "px" && Math.abs(Number(n)) === 1)));
        if (raw.length) findings.push({ at, rule: "length", msg: `${prop}: ${value} → 余白・角丸・文字サイズはトークン（var(--space-*) など）を使う` });
      }
      for (const v of value.matchAll(/var\((--[\w-]+)/g)) {
        if (!defined.has(v[1])) findings.push({ at, rule: "undefined-token", msg: `${v[1]} は tokens.css に無い` });
      }
    }
  }
}

console.log(`token-lint: ${files.length} files, ${findings.length} findings`);
for (const f of findings) console.log(`  ${f.at}  [${f.rule}] ${f.msg}`);
if (strict && findings.length) process.exit(1);
