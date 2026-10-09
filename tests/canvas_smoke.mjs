// 比較キャンバスの動作確認: 選択・コメント・テキスト編集が Copy as prompt の文面に反映されるか。
//   node tests/canvas_smoke.mjs <プロジェクトディレクトリ>
import { createRequire } from "node:module";
import { execSync, spawn } from "node:child_process";
import path from "node:path";

const root = path.resolve(process.argv[2] || ".");
const req = createRequire(path.join(execSync("npm root -g", { encoding: "utf8" }).trim(), "x.js"));
let chromium;
try { ({ chromium } = createRequire(path.join(process.cwd(), "x.js"))("playwright")); } catch { ({ chromium } = req("playwright")); }

const port = 4300 + Math.floor(Math.random() * 500);
const server = spawn("python3", ["-m", "http.server", String(port), "--bind", "127.0.0.1", "--directory", root], { stdio: "ignore" });
await new Promise((r) => setTimeout(r, 800));
const base = `http://127.0.0.1:${port}`;
const assert = (cond, msg) => { if (!cond) { throw new Error(msg); } };

const browser = await chromium.launch();
try {
  // ---- 資料トラック
  {
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
    const errors = [];
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(`${base}/decks/sample/index.html`);
    assert((await page.locator(".card").count()) === 2, "資料: カードが 2 枚でない");
    await page.locator('.card[data-id="B"] input[type=radio]').check();
    await page.locator('.card[data-id="B"] textarea').fill("3枚目に比較表を足す");
    await page.locator("#global-comment").fill("全体で10枚以内");
    await page.locator("#edit-toggle").check();
    const t = page.locator('.card[data-id="B"] .stitle').first();
    await t.evaluate((n) => { n.textContent = "日報入力に25分かかっている"; });
    await page.locator("#show-btn").click();
    const prompt = await page.locator("#prompt-out").inputValue();
    for (const s of ["/deck-lab sample", "採用: 骨子案B「課題起点」", "全体で10枚以内", "骨子案B: 3枚目に比較表を足す", "「日報に25分かかっている」→「日報入力に25分かかっている」"]) {
      assert(prompt.includes(s), `資料: プロンプトに「${s}」が無い\n${prompt}`);
    }
    assert(!errors.length, `資料: JS エラー ${errors.join(", ")}`);
    // 375px で横はみ出しが無いこと
    await page.setViewportSize({ width: 375, height: 800 });
    const over = await page.evaluate(() => document.documentElement.scrollWidth - innerWidth);
    assert(over <= 1, `資料: 375px で横に ${over}px はみ出す`);
    console.log("canvas(deck): OK");
    await page.close();
  }
  // ---- UI トラック
  {
    const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
    const errors = [];
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(`${base}/designs/sample/index.html`);
    await page.waitForTimeout(800);
    assert((await page.locator("iframe").count()) === 2, "UI: iframe が 2 つでない");
    await page.locator("#edit-toggle").check();
    const frame = page.frameLocator('.card[data-id="A"] iframe');
    await frame.locator(".card-value").first().evaluate((n) => { n.textContent = "130件"; });
    await page.locator('.card[data-id="A"] input[type=radio]').check();
    await page.locator("#show-btn").click();
    const prompt = await page.locator("#prompt-out").inputValue();
    for (const s of ["/design-lab sample", "採用: 案A「一覧重視」", "「128件」→「130件」"]) {
      assert(prompt.includes(s), `UI: プロンプトに「${s}」が無い\n${prompt}`);
    }
    assert(!errors.length, `UI: JS エラー ${errors.join(", ")}`);
    console.log("canvas(ui): OK");
    await page.close();
  }
} finally {
  await browser.close();
  server.kill();
}
