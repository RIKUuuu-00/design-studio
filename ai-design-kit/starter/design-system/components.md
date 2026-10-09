# コンポーネント見本

アートボードはここにある部品の組み合わせで作る。新しい部品が必要になったら、まずここに見本を足してから使う。
値はすべて `tokens.css` の変数を使う（直書きは token-lint で検出される）。

## ボタン

```html
<button class="btn btn-primary">保存する</button>
<button class="btn btn-secondary">キャンセル</button>
<button class="btn btn-danger">削除する</button>
```

```css
.btn { font: var(--weight-bold) var(--text-sm)/1 var(--font-sans); padding: var(--space-2) var(--space-4); border-radius: var(--radius-md); border: 1px solid transparent; cursor: pointer; }
.btn-primary { background: var(--color-primary); color: var(--color-on-primary); }
.btn-primary:hover { background: var(--color-primary-hover); }
.btn-secondary { background: var(--color-surface); color: var(--color-text); border-color: var(--color-border); }
.btn-danger { background: var(--color-danger); color: var(--color-on-primary); }
.btn:focus-visible { outline: 2px solid var(--color-focus); outline-offset: 2px; }
```

- 1画面に primary は 1 つまで
- ラベルは動詞で終える（「保存する」「送信する」）

## 入力欄

```html
<label class="field">
  <span class="field-label">メールアドレス</span>
  <input class="input" type="email" placeholder="name@example.com">
  <span class="field-hint">ログインに使う</span>
</label>
```

```css
.field { display: grid; gap: var(--space-1); }
.field-label { font-size: var(--text-sm); font-weight: var(--weight-bold); }
.field-hint { font-size: var(--text-xs); color: var(--color-text-muted); }
.input { font: var(--text-md) var(--font-sans); padding: var(--space-2) var(--space-3); border: 1px solid var(--color-border); border-radius: var(--radius-md); background: var(--color-surface); color: var(--color-text); }
.input:focus { outline: 2px solid var(--color-focus); outline-offset: 0; }
```

## カード

```html
<section class="card">
  <h3 class="card-title">今月の売上</h3>
  <p class="card-value">¥12,340,000</p>
  <p class="card-note">前月比 +8.2%</p>
</section>
```

```css
.card { background: var(--color-surface); border: 1px solid var(--color-border); border-radius: var(--radius-lg); padding: var(--space-5); box-shadow: var(--shadow-sm); }
.card-title { margin: 0; font-size: var(--text-sm); color: var(--color-text-muted); }
.card-value { margin: var(--space-2) 0 0; font-size: var(--text-2xl); font-weight: var(--weight-bold); }
.card-note { margin: var(--space-1) 0 0; font-size: var(--text-xs); color: var(--color-success); }
```

## 表

```css
.table { width: 100%; border-collapse: collapse; font-size: var(--text-sm); }
.table th { text-align: left; color: var(--color-text-muted); font-weight: var(--weight-bold); padding: var(--space-2) var(--space-3); border-bottom: 1px solid var(--color-border); }
.table td { padding: var(--space-3); border-bottom: 1px solid var(--color-border); }
.table td.num { text-align: right; font-variant-numeric: tabular-nums; }
```

- 数値列は右寄せ、桁区切りあり
- 375px 幅では横スクロールではなくカード型に切り替える

## ページ枠

```css
body { margin: 0; background: var(--color-bg); color: var(--color-text); font: var(--text-md)/var(--leading-normal) var(--font-sans); }
.page { max-width: var(--container-max); margin: 0 auto; padding: var(--space-5) var(--gutter); }
```
