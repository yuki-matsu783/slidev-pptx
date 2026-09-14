# samples

## shapes.pptx

`tests/fixtures/deck/shapes.md` をネイティブ書き出しした PPTX です（9 枚）。

- 1〜7 枚目: PowerPoint の図形 187 種を名前順に格子で並べたもの（`PptShape` の `type` に書ける名前の一覧）
- 8 枚目: 見本。調整値・反転・矢じり・文字の枠の効き方
- 9 枚目: 境界の入力。範囲外や数でない調整値・知らない矢じり・知らない図形（書き出しの記録に `W-SHAPE` が出る分）

作り直すときは、いったん `dist/` に書き出してから写します（書き出しの記録 `*.report.json` はローカルのパスと時刻を含むので置きません）。

```sh
mkdir -p dist
pnpm exec slidev-pptx export tests/fixtures/deck/shapes.md -o dist/shapes.pptx
cp dist/shapes.pptx samples/shapes.pptx
```

PowerPoint の実機では開いていません（未確認）。確かめたのは、OPC 整合チェックの error が 0 であることと、Slidev の画面での見え方だけです。
