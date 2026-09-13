---
status: accepted
date: 2026-09-14
ticket: なし（ブランチ ppt-preset-shapes）
---

# PptShape の形は PowerPoint の図形の定義 XML から生成して描く

`PptShape` は 11 種（`rect` `roundRect` `ellipse` `line` と矢印・菱形・三角・六角形）しか受けず、
`rect` `roundRect` は CSS、残りは手書きの SVG で描いていた。PowerPoint の図形の既定の調整値とは別物の近似で、
矢印の柄の太さや六角形の角の位置が PowerPoint で開くと変わり、吹き出しやフローチャート記号は使えなかった。
PowerPoint の図形 187 種の形は ECMA-376 Part 1 の `presetShapeDefinitions.xml` に数式（ガイドと path）で定義されている。
そこで、この XML から描画に要る部分を `src/shapes/presets.ts` に生成し、Slidev の画面ではその数式を枠の実寸で評価して
SVG で描く。書き出しでは名前をそのまま `<a:prstGeom prst>` に出し、PowerPoint は同じ定義で描く。

理由: ネイティブ書き出しの約束は「Slidev で見たものが PowerPoint で直せる形で出る」ことで、図形の形は
縦横比によって変わる（矢印の頭の長さは短辺で決まり、角丸の半径も短辺に比例する）。手書きの SVG でこれを
合わせるには図形ごとに定義の数式を読み解いて写すことになり、187 種では割に合わず、写し間違いも検査で
捕まえにくい。定義の数式は 17 個の演算子と path の 6 命令だけで閉じているので、評価器を 1 つ書けば全種が
同じ規則で描け、調整値（`adj`）も PowerPoint と同じ意味で受けられる。

## Considered Options

| 案 | 退けた理由 |
|---|---|
| **手書きの近似 SVG**（今までの 11 種を 1 種ずつ足す） | 得るのは依存が無く、図形ごとに見た目を調整できること。代わりに形は PowerPoint と一致せず（既定の調整値でも縦横比で崩れる）、調整値を受けるには図形ごとに式を写すことになる。種類を足すたびに SVG と検査が増え、187 種には届かない。「使いたい図形が無い」「開くと形が変わる」の 2 つが残る |
| **定義から生成**（採用） | — |

## Consequences

- **得るもの**: 187 種すべてが 1 つの部品（`type` に `prst` 名）で使え、形と縦横比による変形が PowerPoint と同じ。調整値・反転も PowerPoint の図形と 1:1 で、開いた後に調整ハンドルで直せる。PptxGenJS の `ShapeType` に無いコネクタ 9 種や、`ShapeType` で綴りを誤っている `foldedCorner` も、名前をそのまま渡すので出る。
- **失うもの**: 約 540 KB の XML から作った約 200 行（1 行が長い）の生成物 `presets.ts` をリポジトリに抱え、Slidev の画面の描画に部品ごとの数式の評価が入る（重さは測っていない）。PptxGenJS は `<a:avLst>` を `rectRadius` と `angleRange` でしか書けないので、調整値は後処理 `applyShapeAdjust`（native-export.md §3.2 の 4a）で書き直す。後処理が 1 つ増える。
- **失うもの**: 定義どおりでも一致しない箇所が残る。0 割りと `sqrt` の負数は評価器が 0 に寄せる近似で、PowerPoint の挙動は確かめていない。矢じりの大きさと形、`h` を指定しない図形の文字の枠（Slidev は枠全体）は Slidev 側の近似（ppt-components.md §1.3）。
- **失うもの**: 定義の XML 自体に誤記がある。`pie` の文字の枠（`rect` の `t` と `r` の名前が入れ替わっている）は生成時に `scripts/gen-presets.mjs` の `RECT_FIXES` で補正した。元の XML が直っていれば生成が止まり、補正を消すよう促す。PowerPoint がこの誤記をどう描くかは確かめていない。
- **出典とライセンスの扱い**: 数式は ECMA-376 Part 1（Office Open XML File Formats, Fundamentals and Markup Language Reference）の `presetShapeDefinitions.xml` に由来し、著作権は Ecma International にある。取得元は LibreOffice のリポジトリにある同じファイルで、URL を 1 つのコミット（`500a70ba19d9c1207fd9121531950e55a70fd940`、このファイルを最後に変えたもの）に固定する。生成物の先頭に Ecma の著作権表示・取得元の URL・XML の sha256・補正の件数を書き、XML 本体はリポジトリに置かない。配布の条件（再配布や改変の可否）はこの ADR では精査していない。`slidev-addon-pptx` は `private` でないので、パッケージを公開する前に確かめる。
- **生成のし直し方**: 取得元を変えないなら作り直す必要は無い（生成物は手で直さない）。作り直すときは次のとおり。
  1. `scripts/gen-presets.mjs` の `SOURCE_COMMIT` の URL から XML を取得する（例: `node -e "fetch(process.argv[1]).then(r=>r.arrayBuffer()).then(b=>require('fs').writeFileSync('presetShapeDefinitions.xml',Buffer.from(b)))" <URL>`）。置き場はリポジトリの外
  2. `node packages/slidev-addon-pptx/scripts/gen-presets.mjs <XML のパス>` で `src/shapes/presets.ts` を上書きする。出力の sha256 が生成物の先頭の値と同じなら、元は変わっていない
  3. 取得元を進めるときは `SOURCE_COMMIT` を書き換える。補正が要らなくなっていたらスクリプトが止まるので `RECT_FIXES` から消す
  4. `pnpm test`（`tests/shapes/geometry.test.ts` の座標、`tests/build/presets.test.ts` の 187 種）と `pnpm test:e2e`（`tests/e2e/shapes.test.ts`）を通し、`presets.ts` の差分を読む
