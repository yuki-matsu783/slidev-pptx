# slidev-pptx

Slidev で書いたスライド（`slides.md`）を、PowerPoint で文字・表・図形を直せる PPTX として書き出します。
Slidev 標準の PPTX 書き出し（このリポジトリでは**画像書き出し**と呼びます）は各スライドを 1 枚の画像として貼るので、
受け取った人が中身を編集できません。アドオン `slidev-addon-pptx` は、見出し・段落・箇条書き・表・コードをそのまま
PowerPoint の要素にします（**ネイティブ書き出し**）。向きは `slides.md` → PPTX の一方向で、PowerPoint での手直しは Slidev に戻しません。
用語は [CONTEXT.md](CONTEXT.md) にまとめてあります。

## 手順

```sh
pnpm install        # 依存と Chromium（playwright-chromium）が入る
pnpm export:pptx    # slides.md → slides-export.pptx と slides-export.pptx.report.json
```

書き出しの最後に OPC 整合チェック（PowerPoint の「修復」に掛かる原因の点検）が走り、エラーがあれば exit 1 になります。
標準出力と `*.report.json` に、画像に置き換えた要素の一覧と警告が出ます。

画像書き出し（Slidev 標準）も残してあります。**既定の出力名が同じ `slides-export.pptx` なので、続けて打つと上書きします。**
別名を渡してください。

```sh
pnpm export:pptx-image -- --output slides-image.pptx
```

### 前提

- Node 22.18 以降（`engines` の対象。`bin` はビルド無しで TypeScript を直接読みます。22.6〜22.17 でも `--experimental-strip-types` で起動し直すので動きますが、対象外です）
- Slidev 52 以降、`playwright-chromium` 1.50 以降（書き出しは Slidev のサーバを立てて Chromium で描画し、位置と書式を測ります）。開発時の版は Slidev 52.19.1、PptxGenJS 4.0.1、playwright-chromium 1.63
- フォントは `slides.md` の `fonts` で **游ゴシック / Consolas** に揃えてあります。Mac では OS 同梱の 游ゴシック体 / Menlo で代用されるので、
  計測時の行数が Windows の PowerPoint と少し違うことがあります。5% を超えて溢れる枠には縮小率（`normAutofit`）を書きます（数 % の溢れは残ります）
- 出来上がりは Windows の PowerPoint で確認済みです（2026-09-13、13 枚。「修復」は出ず編集できました）。ほかの環境は未確認です

### CLI

```
slidev-pptx export [entry=slides.md]
  --output, -o <path>     既定 ./slides-export.pptx
  --range <1-3,5>         Slidev の range と同じ綴り
  --theme <name>
  --lang <tag>            文字の言語（校正用）。既定 ja-JP
  --wait <ms>             描画後の追加待機。既定 0
  --timeout <ms>          ページ読み込みの上限。既定 30000
  --report <path>         記録の JSON。既定 <output>.report.json
  --no-check              OPC 整合チェックを飛ばす
  --strict                警告が 1 つでもあれば exit 1
  --keep-server           失敗時にサーバとブラウザを閉じない（調査用）
slidev-pptx check <file.pptx>    OPC 整合チェックだけ
```

exit code: 0 成功 / 1 書き出し失敗または OPC エラー（`--strict` なら警告も）/ 2 引数の誤り。

`--strict` は警告 0 を要求します。headmatter に `transition:` があるデッキは `W-TRANSITION` が必ず 1 件出るので、
`--strict` を使うなら `transition` を外すか、CI では `report.json` の `replacements` と `warnings` の件数を見てください。

## 何が出るか

| Slidev の中身 | PPTX |
|---|---|
| 見出し（スライドの最初の `#` / `##`） | タイトルのプレースホルダー |
| 段落・箇条書き（階層つき）・番号つき・太字・斜体・インラインコード・リンク・上付き/下付き | 本文のプレースホルダーか、描画した位置を測って置くテキスト枠（自由配置） |
| 表 | PowerPoint の表（罫線・結合。見出し行は太字と塗りだけで、PowerPoint の見出し行属性にはならない） |
| コードブロック | 等幅（Consolas）のテキスト枠。色付けは捨てる（行番号は出ない） |
| 引用 | 塗りのあるテキスト枠。左の縦線は捨てる |
| 画像（`<img>`、`layout: image-right` の右半分、表紙の背景） | 画像（URL は取得して埋め込む。SVG は撮影して PNG） |
| 数式（ブロック）・Mermaid・`<svg>`・その他の HTML と Vue 部品 | **画像への置き換え**（2 倍解像度。編集できない） |
| インライン数式 | 文字に平坦化（警告） |
| クリック（`v-click`） | 全部出した状態で 1 枚 |
| `transition`、影・グラデーション・`text-transform` などの装飾 | 捨てる（記録に出る） |

レイアウトは `cover` `default` `center` `two-cols` がスライドマスターのレイアウトに対応します。それ以外（`section` `quote` `image-right` など）は
プレースホルダー無しのレイアウトになり、要素は全部自由配置のテキスト枠です。

## PPT 部品

図形や座標指定が要るところだけ、`Ppt` を接頭辞にした Vue 部品を `slides.md` に書きます（`addons: [slidev-addon-pptx]` が要ります）。
普通の Markdown はそのままネイティブ書き出しされるので、部品は「枠」を足したいときだけ使います。

```md
<PptText :x="60" :y="380" :w="400" fill="#eef" :radius="8" :padding="12" name="lead">

**座標指定**のテキスト枠。`x` `y` を書かなければ、描画した位置と大きさを測って置きます（実測配置）。

</PptText>

<PptShape type="rightArrow" :x="500" :y="400" :w="120" :h="40" fill="#3b82f6" line="none" color="#ffffff">次へ</PptShape>
<PptShape type="line" :x="60" :y="430" :w="620" :h="0" :line="{ color: '#94a3b8', width: 2, tail: 'arrow' }" />
<PptImage src="/photo.jpg" :x="640" :y="120" :w="280" :h="200" fit="cover" alt="写真" />
<PptTable :rows="[['名前','用途'],['default','通常'],['cover','表紙']]" />
```

| 部品 | 主な props |
|---|---|
| 共通 | `x` `y` `w` `h`（Slidev のキャンバス px。1 つでも書けば座標指定）、`export="image"`（画像への置き換え）、`name`（PowerPoint の図形名） |
| `PptText` | `align` `valign`（`h` を指定したときだけ）`size` `color` `bold` `italic` `fill` `line` `radius` `padding`。中身は Markdown（タグの前後に空行） |
| `PptShape` | `type`（`rect` `roundRect` `ellipse` `line` `rightArrow` `leftArrow` `upArrow` `downArrow` `diamond` `triangle` `hexagon`）、`fill` `line` `radius` `rotate` `padding` `align` `valign` `size` `color`。中身は図形の中の文字（`line` では描かない）。形は PowerPoint の図形の近似 |
| `PptImage` | `src` `alt` `fit`（`contain` `cover` `fill`） |
| `PptTable` | `rows`（配列）か中身の Markdown 表、`header`（`rows` のときだけ）`colW` `border` `fill` `size` `align` `valign` |

数式やコードブロックを画像で出したいときは `<div data-ppt-export="image">` で囲みます。

## 記録と警告

`slides-export.pptx.report.json`（と標準出力）に次が出ます。

- `replacements`: 画像に置き換えた要素（スライド番号、理由 `math` / `mermaid` / `svg` / `unknown-element` / `explicit` / `gradient`、撮った大きさ）
- `warnings`: 対処できるもの
  - `W-CSS` 再現できない装飾や色を捨てた / `W-MATH-INLINE` インライン数式を文字にした / `W-LI-BLOCK` 箇条書きの中の表や画像を無視した
  - `W-LAYOUT` 対応表に無いレイアウト、またはプレースホルダーに入れられず自由配置にした / `W-OVERFLOW` 枠に収まらないので縮小率を書いた / `W-OFFSLIDE` スライドの外に掛かるので寄せた
  - `W-IMAGE` 画像を取得できず灰色の矩形にした / `W-LINK` 相対リンクや範囲外のスライドへのリンクを外した
  - `W-HIDDEN` 透明（opacity 0）かスライドの外の要素を飛ばした / `W-INLINE` 段落の中の svg などを無視した
  - `W-NESTED-PPT` `W-PPT-CONTENT` 部品の入れ子や部品の中の表 / `W-DARK` ダーク固定のデッキ / `W-TRANSITION`
  - `W-RENDER` Slidev の描画エラー、ブラウザのエラー、または UnoCSS のクラスが効かないページを測った（位置と色が Slidev と違うかもしれない）
- `dropped`: 規則として捨てたものの件数。`code-highlight`（コードの色付け）、`blockquote-border`（引用の縦線）、`transition`、`background-crop`（背景画像の切り取り）、
  `background-image`（グラデーション）、`console-noise`（Slidev の既知の console エラー）。警告にはしない
- `check`: OPC 整合チェックの結果
- `slideMap`: `--range` で絞ったときの、PPTX の中の番号 → 元のスライド番号（`warnings` と `replacements` の番号は元の番号）

## 検査

```sh
pnpm test        # 単体（ブラウザ無し。後処理・OPC 整合チェック・座標の換算・レイアウト対応表・変換）
pnpm test:e2e    # Slidev のサーバと Chromium を使う（収集器・書き出しの全経路・CLI）
```

置き場と前提は [tests/README.md](tests/README.md)。設計は [wip/design/native-export.md](wip/design/native-export.md) と
[wip/design/ppt-components.md](wip/design/ppt-components.md)、決定の記録は [docs/adr/](docs/adr/)。

## 今回やらないこと

グラフ、SmartArt 相当、クリックのアニメーション変換、既存テンプレートへのマスター差し替え、双方向の同期。
