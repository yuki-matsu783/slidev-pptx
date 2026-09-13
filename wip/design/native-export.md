# ネイティブ書き出しの設計（i0001-02）

対象: `slides.md` → PowerPoint で編集できる PPTX。向きは一方向。
前提の版: PptxGenJS 4.0.1、jszip 3.10.2、Slidev 52.19.1、playwright-chromium 1.63。
調査記録 `wip/research/summary.md` の「設計に効く制約」に従う。制約に反する箇所が出たら
制約の側を再調査して直す（§10 に 1 件ある）。用語は `CONTEXT.md` に従う。部品側の設計は `ppt-components.md`。

判断ごとに「得るもの / 失うもの」を併記する。実装前に確かめる項目は末尾の「要確認」に集める。
本文の「実測」は、この設計を書く過程で PptxGenJS を実際に動かして出力 XML を見た結果と、実装 i0001-04 で確かめた結果を指す。

改訂（i0001-06）: 実装 i0001-04 で実物と食い違った箇所と、設計に無かった判断を書き戻した（§1.2、§3.2〜§3.4、§4.1〜§4.5、§5.2、§7、§8、§11）。
検査は `tests/README.md` にある。

---

## 1. 処理の流れ

```
pnpm export:pptx
  └ slidev-pptx export slides.md [--output out.pptx] ...
      1. resolveOptions({ entry, theme }, 'export')     … options.data に slides / config、options.utils.getLayouts()
      2. createServer(options, { server: { port } })     … /print ルートが有効になる
      3. Playwright: 1 つの context (deviceScaleFactor: 2, colorScheme: light)
         page.goto('/print?print=true&range=...') → 描画完了を待つ（§1.2）
      4. 収集: page.evaluate(collector) → Capture（JSON、§2）
         画像への置き換え: 印を付けた要素を locator.screenshot() で撮る（2 倍）
         画像の取得: <img src> と背景画像を Node 側で fetch して data URL に
      5. 生成: Capture + options.data → PptxGenJS（マスター定義 §5、要素の変換 §4）→ write('nodebuffer')
         生成時に「スライドごとの図形名の列」を PatchContext に記録する（§3.2 の変換 1 が使う）
      6. 後処理: jszip で開く → 変換の列（§3）→ DEFLATE で再圧縮
      7. OPC 整合チェック（§3.4）。エラーなら exit 1、ファイルは残す
      8. 書き出しの記録（§8）を標準出力と JSON に出す
      9. browser.close() → server.close()（失敗時も finally で閉じる。--keep-server のときだけ残す）
```

### 1.1 分担

| 場所 | 役目 | 触ってよいもの |
|---|---|---|
| ブラウザ（収集） | DOM を歩き、位置・書式・中身を **Slidev キャンバス px** のまま JSON にする。要素の分類（ppt-components.md §2.2）と placeholder の候補（§3.1）は DOM の事実として記録するが、レイアウト名と対応表に依る判断はしない | DOM、computed style、`data-ppt-*` 属性 |
| Node（生成） | 単位の換算、レイアウト名の解決とレイアウト対応表、PptxGenJS の呼び出し、置き換え画像の取得、ノート | Capture、`options.data`、`options.utils`、PptxGenJS |
| Node（後処理） | PptxGenJS が出せない・壊すものを ZIP の XML で直す | ZIP の中身だけ |
| Node（検査） | 書き出したファイルの OPC 整合を点検する | ZIP の中身だけ（読むだけ） |

- 得るもの: 収集は「測って記録する」だけなので、ブラウザ側のコードに PptxGenJS もレイアウト表も持ち込まない。Capture を vitest の入力にすれば、生成と後処理をブラウザ無しで検査できる。
- 失うもの: 同じ内容を 2 つの型（DOM の形と Capture の形）で表すぶん、変換規則の変更は両側に及ぶ。

### 1.2 ブラウザの開き方

- サーバは必ず `mode: 'export'` で立てる（`/print` ルートはこのモードでしか無い。`window.__slidev__` は使えないので頼らない）。
- サーバを立てる間は `NODE_ENV=development` にする（元の値は終わったら戻す）。`test`（vitest）や `production` だと Vite の開発サーバで UnoCSS がデッキ由来のクラスを生成せず、`.pt-12` などが永遠に効かない（実測）。
- URL は `/print?print=true`。`PrintSlide.vue` は `print=true` のとき `createFixedClicks(route, CLICKS_MAX)` で描くので、クリック要素は全部出た状態の 1 枚になる（親チケットの決定）。`?print=clicks` はアニメーション変換の段階まで使わない。
  - 得るもの: スライド 1 枚 = PPTX 1 枚で、クリック数を数えなくてよい。
  - 失うもの: `v-click` で「ある段階だけ見せて消す」要素（`v-click.hide` や `[a, b]` の範囲指定）は、最終状態で消えていれば出ない。出したければ `?print=clicks` を使うアニメーション変換を待つ。
- `--range` は `&range=` に渡したうえで、**Node 側でも絞る**（Capture の `slides` を範囲で filter）。Slidev の `/print` は `useNav` の初期化時に `query.range` を 1 度読むだけで、この経路では効かないことがある（実測）。viewport の高さは範囲の枚数で決める。
- 待機は Slidev `exportSlides` の `go()` と同じ手順を自前で持つ: `[data-slidev-no]` の出現 → `.slidev-slide-loading` の消滅 → `[data-waitfor]` → iframe / mermaid / monaco の描画 → `networkidle`。`--wait <ms>` で追加待機。
- その前に **Vite の依存最適化の再読み込み**を待つ。作業ツリーで初めてサーバを立てると "optimized dependencies changed. reloading" でページが読み直され、収集の最中に来ると `page.evaluate` の実行文脈が壊れる。`.print-slide-container` の枚数が揃うまで 1 秒ずつ最大 10 回待ち、`page.evaluate(collect)` が "Execution context was destroyed" で失敗したら 1 回だけやり直す（実測。e2e が初回だけ失敗する原因だった）。
- 待機列の最後に **UnoCSS の生成の確認**を置く。Vite の開発サーバでは、デッキ由来のユーティリティクラス（`pt-12` `text-3xl` など）の CSS が最初の読み込みで生成されないことがあり、そのまま測ると余白や文字の大きさが誤る。2 つの信号を見て、どちらかが無ければ `page.reload()` して待機列をやり直す（最大 3 回）。3 回で来なければそのまま測り、`W-RENDER`（slide 0）を出す。
  - A: `[data-slidev-no]` の下の要素の class のうち、デッキ由来とみなすもの（`slidev-` `ppt-` `katex` `shiki` などの接頭辞を除き、`-` `:` `/` `[` `]` か数字を含むもの。`.katex` `.mermaid` `pre` `svg` の子孫は除く）に CSS 規則がある割合が **50% 以上**。「1 つでも当たれば ok」では判定できない（KaTeX の同梱 CSS や `text-white` は生成前でも当たる。実測: 生成前 0.03、生成後 0.81〜0.87）。
  - B: `__uno.css` / `__uno_shortcuts.css` の `<style data-vite-dev-id>` の長さの合計が **1000 以上**（実測: 生成前 25 / 2572、生成後 1009 / 6539 以上）。その `<style>` 自体が無ければ（UnoCSS を使わないテーマ）B は ok。
  - 得るもの: 生成が遅れた run で黙って誤った値を測らない。
  - 失うもの: しきい値（50%、1000）は Slidev 52.19 と theme-default の実測で、テーマが変わればずれうる。3 回の再読み込みで数秒延びる。
- viewport は `width: canvasWidth, height: canvasHeight × 枚数`（Slidev と同じ。要素の rect は `.print-slide-container` の rect との差で取るので、スクロール位置に依らない）。
- `emulateMedia({ colorScheme: 'light' })`。デッキが `colorSchema: dark` を固定していれば警告 `W-DARK` を出し、測れた色のまま出す。
- headmatter の `transition`（とスライドの frontmatter の `transition`）は捨てる（PptxGenJS に API が無い）。`Report.dropped['transition']` に 1 を出す。警告 `W-TRANSITION` は**出さない**（§8.2 の決定）。

---

## 2. 収集結果（Capture）の形

ブラウザから Node へ渡す JSON。単位は **Slidev キャンバス px**（原点はスライド左上）、色は `#rrggbb`、透明度は 0–100。
換算（インチ・EMU・pt）は Node 側だけがやる。

チケットの成果物には「レイアウト名・ノート・要素の列」とあるが、**レイアウト名とノートは Capture に入れない**。
どちらも DOM に無く（export モードでは `window.__slidev__` が無い。調査 §3）、Node 側の `options.data.slides[no-1]` から取れる。
`export` 指定も Capture には残らず、収集時に `kind: 'image'` と `reason: 'explicit'` に解決される。
- 得るもの: 収集器はブラウザで見えるものだけを記録し、Node 側の情報を二重に持たない。
- 失うもの: Capture だけを見てもスライドのレイアウトが分からない。受入テストの fixture は Capture と `options.data` の相当物を対で持つ。

```ts
interface Capture {
  canvas: { width: number; height: number }        // .print-slide-container の実測（例 980×552）
  slides: SlideCapture[]
}

interface SlideCapture {
  no: number                                        // 1 始まり
  lang: string                                      // [data-slidev-no] の lang 属性。無ければ ''
  zoom: number                                      // .slidev-page の scale。無ければ 1。記録にも出す
  backgroundColor: string                           // .print-slide-container の computed background-color
  elements: Element[]                               // 描画順（DOM 順）
  warnings: Warning[]
}

type Element = TextElement | ShapeElement | ImageElement | TableElement | LineElement

interface ElementBase {
  id: string                    // 'sN-eM'。置き換え画像の撮影に使う data-ppt-capture-id と同じ
  name: string                  // PPT 部品の name prop。無ければ ''。Node 側が一意な名前に確定する（§4.5）
  source: 'markdown' | 'ppt' | 'replaced'
  box: Box                      // 座標指定なら prop の値、実測なら getBoundingClientRect の差分。回転した要素は回転後の外接矩形
  boxSource: 'prop' | 'measured'
  link?: { url: string } | { slide: number }   // 要素全体のハイパーリンク（[![img](a)](b)、PptShape の中の a）
}
interface Box { x: number; y: number; w: number; h: number }

interface TextElement extends ElementBase {
  kind: 'text'
  roleHint?: 'title' | 'body' | 'body2'      // DOM の事実だけ（最初の見出し / その直後の枠 / 右列の最初の枠。ppt-components.md §3.1）。placeholder に入れるかは Node が決める（§4.2）
  paragraphs: Paragraph[]
  frame: FrameStyle             // 枠の塗り・線・角丸・内側余白・透明度
  fit: { contentHeight: number; boxHeight: number }  // scrollHeight と clientHeight。§3.3 の縮小率の材料
  valign: 'top' | 'middle' | 'bottom'
  rotate?: number               // 度。CSS transform の 2D 回転だけ
}
interface Paragraph {
  kind: 'plain' | 'heading' | 'bullet' | 'number' | 'code'
  level: number                 // 見出しなら 1–6、箇条書きなら入れ子の深さ（0 始まり）、それ以外 0
  numberStart?: number          // <ol start>
  align: 'left' | 'center' | 'right' | 'justify'
  lineHeight: number            // computed line-height px
  spaceBefore: number           // computed margin-top px
  spaceAfter: number            // computed margin-bottom px
  runs: Run[]
}
interface Run {
  text: string
  size: number                  // computed font-size px
  color: string
  transparency?: number         // 祖先を含めた opacity の積から。100 - opacity×100
  bold: boolean; italic: boolean; underline: boolean; strike: boolean
  code: boolean                 // 等幅（inline code / kbd / コードブロックの行）
  highlight?: string            // inline code や mark の背景色
  sup?: boolean; sub?: boolean
  link?: { url: string } | { slide: number }
  charSpacing?: number          // letter-spacing px
  breakAfter?: boolean          // <br>
}
interface FrameStyle {
  fill?: { color: string; transparency?: number }
  line?: { color: string; width: number; dash: 'solid' | 'dash' | 'dot' }
  radius?: number               // border-radius px（均一なときだけ）
  inset: [number, number, number, number]   // padding px（上右下左。CSS の順）
  transparency?: number         // opacity
}

interface ShapeElement extends ElementBase {
  kind: 'shape'
  shape: string                 // PowerPoint の図形の名前（ECMA-376 の prst 名。'rect' 'roundRect' 'wedgeRectCallout' ...）。PptShape の type のまま。未知の名前の扱いは Node（§4.3）
  frame: FrameStyle             // roundRect の角丸は frame.radius（px）
  rotate?: number
  paragraphs?: Paragraph[]      // 図形の中の文字
  valign?: 'top' | 'middle' | 'bottom'
  adj?: Record<string, number>  // 調整値。キーは定義の avLst の名前（adj adj1 …）、値は ECMA の単位（例 50000）。data-ppt-opts の adj のうち有限の数値だけ。定義に無い名前の扱いは Node（§4.3）
  flipH?: boolean               // true のときだけ入れる
  flipV?: boolean               // true のときだけ入れる
  arrow?: { head?: string; tail?: string }   // 開いた path の始点（head）・終点（tail）の矢じり。line.head / tail のうち 'none' 以外。値は LineElement と同じ
}
interface LineElement extends ElementBase {
  kind: 'line'                  // <hr> と PptShape type="line"
  line: { color: string; width: number; dash: 'solid' | 'dash' | 'dot'; head?: string; tail?: string }
  from: { x: number; y: number }; to: { x: number; y: number }   // 端点。box は from/to の外接矩形（Node が flipV を決めるのに使う）
}
interface ImageElement extends ElementBase {
  kind: 'image'
  src?: string                  // <img> / 背景画像の絶対 URL または data URL。Node が取得する。SVG なら Node が撮影に切り替える（§4.3）
  captureId?: string            // 置き換え画像。Node が locator.screenshot() で撮る
  reason?: 'math' | 'mermaid' | 'svg' | 'unknown-element' | 'explicit' | 'gradient'
  fit?: 'contain' | 'cover' | 'fill'
  alt?: string
}
interface TableElement extends ElementBase {
  kind: 'table'
  colW: number[]; rowH: number[]          // 実測 px
  headerRows: number
  rows: Cell[][]
}
interface Cell {
  paragraphs: Paragraph[]
  colspan?: number; rowspan?: number
  fill?: string
  border: [BorderSide, BorderSide, BorderSide, BorderSide]   // 上右下左。線が無い辺は { width: 0, color: '' } で必ず埋める
  inset: [number, number, number, number]                    // padding px（上右下左）
  align: Paragraph['align']; valign: 'top' | 'middle' | 'bottom'
}
interface BorderSide { color: string; width: number }

interface Warning {
  code: string                  // §8.2 の一覧
  elementId?: string
  message: string               // 人が読む文。何を出せなかったか、どうすれば出せるか
}
```

- 得るもの: 型がそのまま受入テストの fixture になる。書式は run 単位で平坦なので、PptxGenJS の `TextProps[]` に 1:1 で変換できる。
- 失うもの: 入れ子の HTML（span の中の span）を平坦化するとき、書式の継承を収集側で解く必要がある。computed style を run ごとに読むことで解決するが、要素数ぶん `getComputedStyle` を呼ぶので大きなデッキでは遅い（数百要素で 1 秒程度の見込み。要確認）。

---

## 3. 後処理の層

### 3.1 形

```ts
interface ZipView {                          // jszip の薄い包み
  list(): string[]
  has(path: string): boolean
  readXml(path: string): Document            // @xmldom/xmldom で parse
  writeXml(path: string, doc: Document): void
  readText / writeText / readBinary / writeBinary / remove
}
interface Patch { name: string; run(zip: ZipView, ctx: PatchContext): void }
interface PatchContext {
  capture: Capture
  shapeNames: Record<number, string[]>       // スライド番号 → 生成時に addText/addImage/... した順の図形名（§4.5）
  autofit: Record<number, Record<string, { fontScale: number; lnSpcReduction: number }>>   // §3.3 の結果。図形名で引く
  adjust: Record<number, Record<string, Record<string, number>>>   // スライド番号 → 図形名 → 図形の調整値（avLst の名前 → ECMA の単位の整数）。§4.3 が積み、後処理 4a が書く
  report: Report
}

// パス指定の XML 変換を Patch にする道具
function xmlPatch(name: string, pattern: RegExp, fn: (doc: Document, path: string, ctx: PatchContext) => void): Patch

async function postProcess(buf: Buffer, patches: Patch[], ctx): Promise<Buffer>
//  loadAsync → 順に run → generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
```

XML は文字列置換ではなく DOM（`@xmldom/xmldom`）で触る。

- 得るもの: id の振り直しや図形の削除は構造を見ないとできない。正規表現で XML を触ると属性順や名前空間接頭辞の違いで壊れる。
- 失うもの: 依存が 1 つ増える。シリアライズで名前空間宣言や空要素の書き方（`<a:p/>` と `<a:p></a:p>`）が変わりうるので、「何もしない Patch を通した結果が元と同値」を受入テストで固定する。

### 3.2 変換の列（順番は固定）

| # | 名前 | 対象 | すること | 必須 |
|---|---|---|---|---|
| 1 | `renameShapes` | `ppt/slides/*.xml` `ppt/slideLayouts/*.xml` `ppt/slideMasters/*.xml` | `<p:spTree>` の子（`sp` `pic` `graphicFrame` `cxnSp` `grpSp`）を文書順に見て、`<p:cNvPr id>` を 2 から振り直す（1 は `<p:nvGrpSpPr>` 用に空ける）。スライドでは `name` を `ctx.shapeNames[no][index]` で**付け直す**。列より多い図形（書き出し時に自動追加された空 placeholder）は `Placeholder N` | **必須** |
| 2 | `dropEmptyPlaceholders` | `ppt/slides/*.xml` | `<p:ph>` を持つ `<p:sp>` のうち `<a:t>` に文字が無いものを消す | 必須（`masterName` 付きスライドには未使用 placeholder が必ず足される。実測） |
| 3 | `dedupeParagraphProps` | `ppt/slides/*.xml` `ppt/notesSlides/*.xml` | 各 `<a:p>` の `<a:pPr>` を先頭の 1 つだけ残す（2 つ目以降を消す。先頭の `<a:pPr>` が `<a:r>` の後ろにあれば最初の子に移す） | **必須**（PptxGenJS は run ごとに `<a:pPr>` を出し、`bullet` の無い run にも `<a:buNone/>` 付きの `<a:pPr>` を書く。太字やリンクを含む段落は必ず 2 つ以上になる。実測。`CT_TextParagraph` は先頭に高々 1 つ） |
| 4 | `applyAutofitScale` | `ppt/slides/*.xml` | `ctx.autofit[no]` の図形名で `<p:sp>` を引き、`<a:bodyPr>` の**既存の** `<a:normAutofit>` に `fontScale` `lnSpcReduction` を足す（無ければ 1 つ作る。2 つにはしない） | 条件つき |
| 4a | `applyShapeAdjust` | `ppt/slides/*.xml` | `ctx.adjust[no]` の図形名で `<p:sp>` を引き（`p:cxnSp` `p:pic` などは触らない）、`<a:prstGeom>` の `<a:avLst>` の中身を `<a:gd name="…" fmla="val N"/>` で**置き換える**（無ければ作る）。並びは定義（`src/shapes/presets.ts`）の avLst の順、値は整数に丸める。定義に無い名前と、`prst` が定義に無い図形は書かない。`adjust` が空のスライドは触らない（既存の gd が残る） | 条件つき（PptxGenJS は avLst を `rectRadius` と `angleRange` でしか書けない。§4.3） |
| 5 | `splitNotesParagraphs` | `ppt/notesSlides/*.xml` | `<a:t>` の改行を `<a:p>` の区切りに割る。区切りは `\r\n` `\r` `\n` のどれでも（PptxGenJS は CRLF で詰めるが、`@xmldom/xmldom` は XML の行末正規化で parse 時に LF にするので、DOM で見えるのは LF。実測）。割って作る `<a:p>` には元の `<a:pPr>` を先頭に 1 つだけ複製する（3 の後に走るので、C14 の不変条件を自分で守る） | 必須（PptxGenJS は 1 つの `<a:t>` に CRLF で詰める。実測） |
| 6 | `replaceMaster` | マスター・レイアウト・テーマ・各スライドの rels | **将来のマスター差し替えの継ぎ目**。今回は no-op。§5.4 | 継ぎ目 |
| 7 | `rebuildContentTypes` | `[Content_Types].xml` | ZIP の実在パートから作り直す。`Default` の**拡張子の集合は ZIP に実在するもの**から（PptxGenJS の出力を写さない）、拡張子 → ContentType の対応は固定表で持つ（`jpg` / `jpeg` → `image/jpeg`、`webp` → `image/webp`。PptxGenJS 自身は `jpg` → `image/jpg` を出すが IANA の型に揃える。表に無い拡張子は `application/octet-stream`）。`Override` はパートの種類ごと | **必須**（実在しない `slideMasterN.xml` の Override が混ざる。実測） |

変換 1 が図形の名前を確定させる理由: PptxGenJS は placeholder 指定の `addText` で、呼び出し側の `objectName` をレイアウト側の名前（`Text 0`）で上書きする（`pptxgen.cjs.js` L2537。実測でタイトル枠が `name="Text 0"` になり、同じスライドの自動追加 body も `Text 0` で重複した）。つまり `objectName` は自由配置の図形にしか効かない。そこで名前は API に頼らず、**生成時に記録した順序**で後処理が付け直す。`_slideObjects` は add 呼び出しの順で書かれ、自動追加 placeholder はその後ろに付くので、順序は Node 側で確定できる。

`ppt/notesSlides/*.xml` は変換 1 の対象外。PptxGenJS はノートの `cNvPr id` を一意に出す（実測で 1〜4）。C6 はノートも検査するが保険で、error が出たら変換 1 の対象に足す。

- 順番は 1 → 7 で固定（4a は 4 の直後）。名前で図形を引く変換（4、4a、将来のアニメーション）は 1 の後。7 は最後。4a に番号を振り直さず枝番にしたのは、既存の検査とコードのコメントが「後処理 5」などの番号で変換を指しているため。将来の追加（グループ化 `<p:grpSp>`、`<p:timing>`、`<p:transition>`、`<a:latin>` の書き分け）は 3 と 6 の間。
- 得るもの: 各変換は前の変換の結果だけを前提にすればよい。
- 失うもの: 変換を足すたびに位置を決める必要があり、並びが暗黙の依存になる。各変換の先頭コメントに「前提にする変換」を書いて依存を明示する。

### 3.3 `normAutofit` の縮小率

- 全てのテキスト枠（placeholder と自由配置）に `fit: 'shrink'` を付けて `<a:normAutofit/>` を出す（placeholder 指定でも出ることは実測済み）。PowerPoint は編集時に自分で縮小率を計算し直す。
- 書き出し直後に溢れないよう、2 つの条件のそれぞれで「枠 ÷ 内容」の比を出し、**小さいほう**を採る:
  - A: 収集した `fit.contentHeight > fit.boxHeight`（Slidev 側で既に溢れて隠れている）→ `boxHeight / contentHeight`
  - B: 出す枠の高さより `fit.contentHeight` が大きい → `枠の高さ / contentHeight`。枠の高さは、placeholder なら対応表（§5.2）の `h`（対応表の px はキャンバス px と同じ空間）、自由配置なら寄せた後の `box.h`
- 比が **0.95 未満**（5% を超えて溢れる）のときだけ、`fontScale = floor(比 × 100) × 1000`（下限 25000）、`lnSpcReduction = 10000`（fontScale < 90000 のとき 20000）を `ctx.autofit` に入れ、警告 `W-OVERFLOW` を記録する。5% の余裕は、計測環境と PowerPoint のフォント差で行数が数 % 動くため（数 % の縮小率は当たらないうえ、警告が雑音になる）。
- 得るもの: 開いた瞬間に文字が枠からはみ出ない。
- 失うもの: 縮小率は Slidev のフォントで測った推定。PowerPoint の游ゴシックでは行数が変わりうるので、正確には合わない。PowerPoint がリサイズ時に計算し直すので、ずれは 1 回の編集で解消する。5% 未満の溢れはそのまま出る（PowerPoint が開くときに `normAutofit` で縮める）。

### 3.4 OPC 整合チェッカ

`slidev-pptx check <file.pptx>` と、書き出しの最後（`--no-check` で飛ばせる）と、vitest から同じ関数を呼ぶ。python-pptx で読めても PowerPoint が「修復」に掛ける項目を狙う。

| 規則 | 内容 | 重さ |
|---|---|---|
| C1 well-formed | 全 `.xml` `.rels` が parse できる | error |
| C2 Override の実在 | `[Content_Types].xml` の各 `Override PartName` が ZIP に実在する | error |
| C3 型の抜け | 全パートが `Default`（拡張子）か `Override` のどちらかで型を持つ | error |
| C4 rels の解決 | 全 `.rels` の `Target`（`TargetMode="External"` 以外）が実在パートを指す | error |
| C5 rels の必須 | 各 `slideN.xml` に `slideLayout` の関係が 1 つ、`presentation.xml.rels` に `slideMaster` が 1 つ以上 | error |
| C6 id の一意 | スライド・レイアウト・マスター・ノートの各パート内で `cNvPr id` が一意で、0 でない | error |
| C7a sldId の解決 | `presentation.xml` の `sldId r:id` / `sldMasterId r:id` / `notesMasterId r:id` が `presentation.xml.rels` に在る | error |
| C7b sldLayoutId の解決 | `slideMasterN.xml` の `sldLayoutId r:id` が `slideMasterN.xml.rels` に在る | error |
| C8 placeholder の対応 | スライドの `<p:ph idx type>` が、そのレイアウトに同じ `idx` の placeholder を持つ（`idx` 無しの `type="title"` は可） | warn |
| C9 xfrm の値 | `<a:ext cx="0" cy="0"/>` の図形（PowerPoint は開けるが選べない）は warn。`<a:off>` `<a:ext>` の `x y cx cy` が負なら error（§4.1 の負のインチ。実測） | warn / error |
| C10 r:* の逆引き | 各パートの `r` 名前空間（relationships）の全属性（`r:id` `r:embed` `r:link`、将来のグラフの `r:dm` など）の値が、そのパートの rels に在る | error |
| C11 autofit の重複 | `<a:bodyPr>` の子に `normAutofit` `spAutoFit` `noAutofit` が 2 つ以上ない | error |
| C12 ph idx の一意 | 1 スライド内で `idx` 属性を持つ `<p:ph>` の `idx` が重複しない（`idx` 無しは C8 の扱い） | error |
| C13 不正文字 | 全テキストノードと全属性値が XML 1.0 で許される文字だけ（`<a:t>` のほか `cNvPr name` `descr` `tooltip` `cSld name` にも利用者の文字列が入る。PptxGenJS は `& < > " '` しかエスケープせず、`@xmldom/xmldom` は制御文字を通すので C1 では止まらない）。生成時に `sanitize.ts` で取り除くのが本線で、この検査は保険。`sanitize.ts` と C13 は XML 1.0 で合法な DEL / C1（0x7F–0x9F）も対象にする（見えない文字で、PowerPoint の表示で害になりうる） | error |
| C14 pPr の重複 | `<a:p>` 直下の `<a:pPr>` は高々 1 つで、あれば最初の子（後処理 3 の結果を固定する） | error |

出力は `CheckResult[]`（`src/types.ts`。`{ rule, part, message, level: 'error' | 'warn' }`）。error が 1 つでもあれば exit 1。
- 得るもの: 「修復」の原因を CI で止められる。LibreOffice の無い環境でも動く。
- 失うもの: PowerPoint の修復条件は公開されていないので、規則は経験則。実機で修復が出たら規則を足す運用になる。

---

## 4. 要素の変換（Capture → PptxGenJS）

### 4.1 座標の換算

スライドは `LAYOUT_WIDE`（12192000 × 6858000 EMU、13.333 × 7.5 インチ。`pptxgen.cjs.js` L7090）。
換算は **幅だけ**で決める。

```
SLIDE_W_EMU = 12192000
SLIDE_W_IN  = SLIDE_W_EMU / 914400            // 13.3333…。定数はこの 1 つから導く
emuPerPx = SLIDE_W_EMU / canvas.width         // 980 px なら 12440.816…
emu(px)  = Math.round(px * emuPerPx)
pt(px)   = Math.round(px * 72 * SLIDE_W_IN / canvas.width * 10) / 10   // 980 px なら px × 0.9796、小数 1 桁
inch(px) = px * SLIDE_W_IN / canvas.width                              // rectRadius 用（§4.2）
```

**座標（x y w h、表の colW rowH）は EMU の整数で渡す。** PptxGenJS は `x y w h` を `getSmartParseNumber`（100 未満ならインチ、**100 以上**なら EMU）で、表の `colW` `rowH` を `inch2Emu`（**100 より大きい**ときだけ EMU。ちょうど 100 は 100 インチになる。実測）で解く。
- 0 は 0 のまま。1 以上 100 以下は **101** に切り上げる（座標も表も同じ関数で。0.001 インチ未満なので見えない）。
- **負の値**はインチ扱いになって壊れる（`-5` → −5 インチ。PptxGenJS に負のインチを渡すと `cx="-113758675200"` のような値が出る。実測）。スライドの外に掛かる要素は**常に**内側に寄せる: x / y が負なら 0 に、右端・下端がスライドを超えていれば w / h を縮める（左端・上端は動かさない）。縮めた結果 w か h が 0 以下（x が −100 で w が 50 など、全体が外にある）なら要素を除外して `W-HIDDEN`。placeholder に入る枠は座標を渡さないので寄せない。
- 線（`hr` と `PptShape type="line"`）は `box` でなく**端点**（`from` / `to`）をキャンバスの中に寄せる（両端をそれぞれ 0〜幅、0〜高さに丸める。傾いた線は向きを保つ）。どちらかの軸で全体が外（水平線が y < 0、縦線が x > 幅、など）なら除外して `W-HIDDEN`。§3.4 の C9 は `cx cy` に加えて負の `x y cx cy` も error にする。
- 寄せた量が **5 px 以上**のときだけ `W-OFFSLIDE` を出す。Slidev の `h1 { -ml-[0.05em] }` で x が −2 px になるのは全スライドで起きるので、5 px 未満は寄せるだけで黙る。
- フォント pt は PptxGenJS が `sz = round(pt × 100)` にする。

- キャンバス高さは client 側 552 px（ceil）と export 側 551 px（round）でずれる。幅だけで換算し、高さは 7.5 インチ固定にすると、552 px は 7.51 インチに相当して 0.01 インチ（0.7 pt）はみ出る。無視する。
- 得るもの: 縦横比が保たれ、換算が 1 つの定数で済む。
- 失うもの: 最下端に置いた要素が 0.7 pt だけスライドの外に掛かる。目には見えない。

### 4.2 テキスト枠

`slide.addText(runs, opts)`。`runs` は `Paragraph[]` を `TextProps[]` に平坦化したもの。段落の切れ目は最後の run に `breakLine: true`、段落書式（`bullet` `indentLevel` `align` `paraSpaceBefore/After` `lineSpacingMultiple`）は段落の先頭 run に付ける。

| Capture | PptxGenJS | 単位 |
|---|---|---|
| `roleHint` があり、§5.1 で解いたレイアウトの対応表（§5.2）に同名の placeholder がある | `placeholder: roleHint`（座標は渡さない。レイアウトの値が勝つ）。**role の確定は Node の仕事**。収集器はレイアウト名も対応表も知らないので、候補（`roleHint`）だけを記録する | |
| `roleHint` があるが対応表に無い（`blank`、`two-cols` 以外の `body2` など） | 自由配置にする。`blank` 以外なら `W-LAYOUT` に理由を添える（綴り違いの placeholder 名を PptxGenJS に渡すと `<p:ph>` の無い枠が左上に出る。実測） | EMU |
| 同じスライドで同じ `roleHint` の枠が 2 つ以上 | 文書順で最初の 1 つだけ placeholder に入れ、2 つ目以降は自由配置にして `W-LAYOUT`。同じ role を 2 つ入れると `<p:ph idx>` が重複して C12 で止まる（実測）。今の収集器は `title` 候補を 1 スライドに 1 つしか出さないので、これは収集器の変更や手書きの Capture に対する保険 | EMU |
| それ以外（自由配置） | `x y w h` | EMU |
| `rotate` | `rotate` | 度 |
| `Paragraph.kind: 'bullet'` | `bullet: { characterCode: '25AA' }`（Slidev の square）、`indentLevel: level`。字下げ幅は PptxGenJS の 1 段 342900 EMU（約 27 px）に任せる（Slidev の `li { ml-1.1em pl-0.2em }` ≈ 23 px と近い）。`bullet.indent` は記号と文字の間隔で、字下げではないので使わない | |
| `'number'` | `bullet: { type: 'number', numberStartAt: numberStart }`、`indentLevel: level`。番号の種類は既定（`arabicPeriod`）に任せる（`numberType` は実装が読まず、`style` は deprecated。`numberStartAt` は `startAt` に出る。実測） | |
| `'heading'` | 段落書式は plain と同じ。大きさは run の `size` で出る | |
| `'code'` | run に `fontFace: <等幅>`（§6）。1 行 1 段落 | |
| `align` | `align` | |
| `lineHeight / size` | `lineSpacingMultiple`（1.0 未満は 1.0） | |
| `spaceBefore` `spaceAfter` | `paraSpaceBefore` `paraSpaceAfter` | pt |
| `Run.link.url` | `hyperlink: { url }`。`{ slide }` → `hyperlink: { slide }` | |
| `Run.code` | `fontFace: <等幅>`、`highlight: <inline code の背景色>` | |
| `Run.sup / sub` | `superscript` / `subscript` | |
| `Run.transparency` | `transparency`（整数に丸める） | 0–100 |
| `Run.charSpacing` | `charSpacing` | pt |
| `frame.fill` `line` | `fill: { color, transparency }` `line: { color, width, dashType }`。Capture の `dash: 'dot'` は PptxGenJS の型に無いので `'sysDot'` に変換する | `width` は pt |
| `frame.radius` | `rectRadius` と `shape: 'roundRect'` | **インチ**（`inch(px)`。PptxGenJS は `adj = round(rectRadius × 914400 × 100000 / min(cx, cy))` で、EMU を渡すと桁あふれする） |
| `frame.inset`（上右下左） | `margin: [左, 右, 下, 上]` | pt。**PptxGenJS 4.0.1 の実装は `[l, r, b, t]`**（`pptxgen.cjs.js` L5389-5392）。型定義のコメントは TRBL と書いてあるが実装と食い違う。並べ替えは `units.ts` の 1 関数に閉じる |
| `frame.transparency` | `fill.transparency` にだけ反映する。run の `transparency` は収集時に祖先の opacity を掛け込んであるので、ここで重ねると二重掛けになる | |
| `valign` | `valign` | |
| 全枠 | `fit: 'shrink'`、`fontFace: <本文フォント>`、`lang: <§6>`、`isTextBox: true`（placeholder 以外）。`objectName` は渡さない（§3.2 変換 1 が付ける） | |

### 4.3 図形・線・画像・表

| 種類 | 呼び出し |
|---|---|
| shape | 文字があれば `addText(paragraphs, { shape, fill, line, rotate, flipH, flipV, x y w h })`、無ければ `addShape(shape, { fill, line, rotate, flipH, flipV, x y w h, hyperlink })`。`shape` は下の「図形の名前と調整値」で決める。`shape` と `fill` の同時指定は実測で `<a:prstGeom>` と塗りが両方出る。**要素リンクを `addText` に渡さない**: `addText` の `hyperlink` は run の分しか rels に登録せず、`<a:hlinkClick r:id="rIdundefined">` が出る（実測。C10 で必ず落ちる）。文字のある図形と、`a` で囲まれたテキスト枠の `link` は全 run の `hyperlink` に付ける（下線つきになる） |
| line | `addShape(ShapeType.line, { x y w h, line: { color, width, dashType, beginArrowType, endArrowType }, flipV })`。`from/to` から `x y w h` と `flipV` を決める（`PptShape` の `flipH` `flipV` prop は線では無視する）。`head` `tail` の値が `none` `arrow` `diamond` `oval` `stealth` `triangle` に無ければ `arrow` にして `W-SHAPE`（図形と同じ） |
| image | `addImage({ data: 'data:image/png;base64,…', x y w h, altText, sizing, hyperlink })`。`fit: 'contain' \| 'cover'` → `sizing: { type, w, h }`、`'fill'` と未指定 → `sizing` 無し（`type: 'fill'` を渡すと `addImage` は通るが `write()` が `TypeError` で落ちる。実測）。`addImage` の `hyperlink` は自前で rels を登録するので使える。URL は Node が fetch し、失敗したら `W-IMAGE` を出して灰色の矩形（`rect`）を置く |
| image（SVG） | `src` の拡張子または `content-type` が SVG なら `addImage` に渡さず、その `<img>` を撮影に回す（`data-ppt-capture-id` を振って `locator.screenshot()`）。PptxGenJS の SVG 経路は `image-N.png` の中身に SVG を書くので壊れた PPTX になる（`addImageDefinition` の `isSvgPng`）。置き換え一覧に `reason: 'svg'` で載せる |
| table | `addTable(rows, { x y w h, colW: EMU[], rowH: EMU[] })`。表の高さは `rowH` の和で決まるので、寄せた枠（`box.h`）より高いときは各 `rowH` を比例で縮め、`h` には縮めた後の和を渡す。`headerRows` は**渡さない**（PptxGenJS に `<a:tblPr firstRow>` を出す口が無い。見出し行は Capture の太さ・塗りで表す。実測）。セルは `{ text: TextProps[], options: { colspan, rowspan, fill, align, valign, border, margin } }`。`border` は **4 辺必ず埋める**（未指定の辺は PptxGenJS が `DEF_CELL_BORDER`（solid / 666666 / 1 pt）で補うので、線の無い辺を明示しないと罫線が増える。`undefined` があっても落ちはしない。実測）: `width: 0` → `{ type: 'none' }`、それ以外 → `{ type: 'solid', pt: pt(width), color }`。順序は **上右下左**（テキスト枠の `margin` と違う。実測）。`margin` は上右下左の pt。**上の値が 1 未満だと 4 辺ともインチ扱いになる**（PptxGenJS の `cellMargin[0] >= 1` の分岐）ので、上は 1 pt に切り上げる。空白の `hMerge/vMerge` は PptxGenJS が作る |
| background | `slide.background = { color: backgroundColor }`。§5.5 の背景画像があれば `{ data }` |

**図形の名前と調整値**（`convert.ts` の `addShape`）:
- `shape` が PowerPoint の図形の定義（`src/shapes/presets.ts` の 187 種。`isPreset`）にあれば、名前を PptxGenJS の `ShapeType` を通さずそのまま渡す（PptxGenJS は `'<a:prstGeom prst="' + shape + '">'` と書く）。`ShapeType` に無いコネクタ 9 種と、`ShapeType` で綴りを誤っている `foldedCorner` もこれで出る。定義に無ければ `rect` にし、調整値も捨てて `W-SHAPE`。
- 調整値: `adj` を `normalizeAdjust`（`src/shapes/`。Slidev の描画と共通）で定義の avLst の名前に絞り、整数に丸めて `ctx.adjust[PPTX の番号][図形名]` に積む。後処理 4a が `<a:avLst>` に書く。定義に無い名前・数でない値・丸めて 32 bit 整数（`ST_Coordinate32`）の範囲外の値は捨てて `W-SHAPE`（範囲外は PowerPoint が「修復」に掛ける）。
- `roundRect` の角丸: `frame.radius`（px）を `rectRadius` に渡さず、`adj.adj` に換算する（角の半径 = 短辺 × `adj` / 100000。`PptShape.vue` と同じ式）。上の絞り込みで捨てられなかった `adj.adj`（有限で、丸めて 32 bit 整数に収まる）があればそちらが勝つ（`PptShape.vue` も同じ判定）。PptxGenJS は `rectRadius` が 0 だと捨てて（`if (rectRadius)`）PowerPoint の既定の角丸になるので、0 を保つためにこちらにした。短辺は**寄せた後の枠**（§4.1）で取り、px の半径を保つ。値は定義の `pin 0 adj 50000` に合わせて 0〜50000 に収める。テキスト枠（`PptText` などの `frame.radius`）は今までどおり `rectRadius`（§4.2）。
- 反転: `flipH` `flipV` は `true` のときだけオプションに渡す（`<a:xfrm flipH flipV>`）。
- 矢じり: `arrow.head` → `line.beginArrowType`、`arrow.tail` → `line.endArrowType`。値は `none` `arrow` `diamond` `oval` `stealth` `triangle` のどれかで、それ以外は `arrow` にして `W-SHAPE`。
- 文字の枠は PPTX に書かない。PowerPoint は定義の文字の枠を自分で使う（Slidev 側の扱いは ppt-components.md §1.3）。
- 得るもの: PowerPoint の図形 187 種が名前 1 つで出て、調整値・反転も PowerPoint の図形と 1:1。
- 失うもの: 調整値は枠に対する比なので、スライドからはみ出して枠を縮めた図形では、形が Slidev の見た目とずれる（直していない）。`roundRect` の角の半径だけは寄せた後の枠で換算して px を保つので、他の図形と扱いが揃っていない。

`dash` の変換（Capture の `'dot'` → PptxGenJS の `'sysDot'`、他はそのまま）は、テキスト枠・図形・線で共通に `units.ts` の `toDashType()` を使う。

同じ画像を複数スライドで使うと、PptxGenJS は `path` 一致でしか重複を除かないので、data URL では毎回 `ppt/media` に入る。Node 側は fetch の結果を `src` でキャッシュして取得は 1 回にするが、ZIP には枚数ぶん入る。
- 得るもの: 実装が単純。
- 失うもの: ロゴを全スライドに置くとファイルが大きくなる。困ったら後処理で同じ内容の media を 1 つにまとめる変換を足す。

### 4.4 ノート

`options.data.slides[i].note`（生 Markdown）を `addNotes` に渡す。前処理は 3 つだけ: `[click]` `[click:N]` を後ろの空白ごと消す、行頭の `- ` `* ` を `• ` に（字下げの後でも同じ。字下げは保つ）、`**` `_` `` ` `` の記号をそのまま残す。行末は LF に揃える（CRLF のノートでも後処理 5 の区切りが 1 つになる）。`undefined`（ノート無し）は空文字。
後処理 5 が行ごとに `<a:p>` に割る。
- 得るもの: Markdown の変換器を Node 側に持ち込まない。
- 失うもの: ノートの太字やリンクは記号のまま見える。ノートの run は PptxGenJS が `lang="en-US"` で出し、`--lang` は効かない（`addNotes` に lang の口が無い。実測）。校正の言語を合わせたければ後処理に足す。

### 4.5 図形の名前

Node 側が、スライドごとに図形を add する順に名前を確定し、`ctx.shapeNames[no]` に積む。

- 列の先頭は **Node が足す図形**（§5.5 の `Background dim`。あれば 1 つ）。その後に Capture の要素を add した順。
- `name` prop があればそれ。同じスライドに同名があれば `-2` `-3` を足す。
- 無ければ、placeholder に入れると確定した枠（§4.2）は固定名 `Title` `Body` `Body 2`。それ以外は `<種類> <要素番号>`（`Text 3`、`Shape 4`、`Image 5`、`Table 6`、`Line 7`、`Replaced 8`。要素番号は `ElementBase.id` の M）。
- 列より多い図形（書き出し時に自動追加された空 placeholder）は `Placeholder <spTree の位置>`。後処理 2 が消すので通常は残らない。
- `name` に XML で不正な文字や制御文字があれば、生成時に取り除く（`sanitize.ts`。§3.4 C13 は保険）。
- 後処理 1 がこの列で `<p:cNvPr name>` を付け直す。`W-*` の `elementId` と `Report` にはこの名前も出す。

- 得るもの: 名前が PowerPoint の「選択ウィンドウ」でそのまま見え、警告と突き合わせられる。
- 失うもの: 生成と後処理が「add した順 = spTree の順、自動追加 placeholder は末尾」という PptxGenJS の内部の性質に依存する。版を上げたときに受入テストで確かめる（§11 の 6）。

---

## 5. レイアウト対応表

### 5.1 レイアウト名の解決（Node 側）

Slidev と同じ順に 2 段で解く。

1. Slidev が使うレイアウト名: `frontmatter.layout` → `slides[0].frontmatter.defaults.layout` → 1 枚目なら `cover`、それ以外 `default`。その名前が `await options.utils.getLayouts()` に**無ければ `default`**（Slidev が `default` として描くので、こちらも合わせる。`serve-*.mjs` L683-690）。
2. 対応表: 1 の結果が §5.2 の表に無ければ `blank`（警告 `W-LAYOUT`。これは Node 側の警告なので `Report.warnings` に入る）。

DOM の `.slidev-layout` の class は使わない。`two-cols` の class は `two-columns` で名前が一致せず、`image-right` は内側に `.slidev-layout.default` を持つ（どちらも `client/layouts/` で確認）。

### 5.2 対応表（canvasWidth 980 のときの px と、換算後のインチ）

Slidev の `.slidev-layout` は `px-14 py-10`（左右 56 px、上下 40 px）。内容幅 868 px。1 インチ = 73.5 px。

| Slidev | PptxGenJS `defineSlideMaster.title` | slideLayout 番号 | placeholder | x | y | w | h |
|---|---|---|---|---|---|---|---|
| （組み込み） | `DEFAULT` | 1 | なし | | | | |
| 表にない全て | `blank` | 2 | なし | | | | |
| `cover` | `cover` | 3 | `title`（type title） | 0.76 (56) | 2.99 (220) | 11.81 (868) | 1.09 (80) |
| | | | `body`（type body） | 0.76 | 4.19 (308) | 11.81 | 1.50 (110) |
| `default` | `default` | 4 | `title` | 0.76 | 0.54 (40) | 11.81 | 0.60 (44) |
| | | | `body` | 0.76 | 1.31 (96) | 11.81 | 5.66 (416) |
| `center` | `center` | 5 | `title` | 0.76 | 3.21 (236) | 11.81 | 0.60 |
| | | | `body` | 0.76 | 3.97 (292) | 11.81 | 1.50 |
| `two-cols` | `two-cols` | 6 | `title` | 0.76 | 0.54 | 5.90 (434) | 0.60 |
| | | | `body` | 0.76 | 1.31 | 5.90 | 5.66 |
| | | | `body2` | 6.67 (490) | 0.54 | 5.90 | 6.42 (472) |

根拠:
- `default`: h1（`text-4xl`、行 40 px、`mb-4` 16 px）の直後 96 px から本文。見出しが h2（`text-3xl`、行 36 px、下余白なし）のときは本文が 76 px から始まるので、タイトル枠が 8 px 広く、本文枠が 20 px 下から始まる。実害は枠の余白だけ。
- `cover`: h1 60 px / 行高 80 px + `mb-4` 16 − `h1 + p` の `-mt-2` 8 + 副題 1 行 24 = 112 px のブロックを内容領域 472 px の縦中央に置いた上端 40 + (472 − 112) / 2 = 220 px。副題の上端は 220 + 80 + 16 − 8 = 308 px。
- `center`: h1 36 px 行 40 + 16 + 1 段落 24 = 80 px のブロックを中央に置いた 236 px。
- `two-cols`: `grid-cols-2`、gap 無し。

placeholder の位置は**固定**で、実測しない。PptxGenJS は placeholder 指定時にレイアウトの座標を明示した座標より優先する（`pptxgen.cjs.js` L5151-5159）。
- 得るもの: PowerPoint 側で「レイアウトのリセット」「新しいスライド」が意味を持つ。アウトライン表示とアクセシビリティ検査がタイトルを認識する。
- 失うもの: Slidev で縦中央に置かれた表紙の位置や、`layoutClass: gap-8` の 2 段組の列幅（左 418 px、右は 506 px から）は、表の固定値（434 / 490）から最大 16 px ずれる。位置を Slidev に合わせる必要が出たら、後処理に「placeholder の `<a:xfrm>` を実測値で上書きする」変換を足す（この設計では入れない）。
- 失うもの: `body` placeholder は内容領域いっぱい（`default` で 416 px）に固定なので、その下に置いた自由配置の要素（コード枠・表・画像）と枠が重なる。見た目は空欄で問題ないが、PowerPoint で下の要素を掴みにくい。困る場合は選択ウィンドウから選ぶか、Slidev 側で本文を `<PptText>` にして自由配置にする。
- 表にない `section` `quote` `image-right` `fact` `statement` `intro` `end` などは `blank` で、見出しも本文も自由配置（実測）のテキスト枠になる。

### 5.3 マスターの定義

```ts
// packages/slidev-addon-pptx/src/build/masters.ts   ← マスター定義はこのファイルだけ
export const LAYOUTS: Record<'blank' | 'cover' | 'default' | 'center' | 'two-cols', SlideMasterProps>
export function defineMasters(pptx: PptxGenJS, canvasWidth: number): void   // 表の順で defineSlideMaster を呼ぶ
export function masterFor(layout: string): keyof typeof LAYOUTS             // §5.1 の 2 段目
```

- 共通: `background: { color: 'FFFFFF' }`（スライド側で実測色を上書き）、`margin: 0`（マスター側の属性。`createSlideMaster` は placeholder の options に配らないので、下の「座標以外を置かない」とは衝突しない。実測）、`slideNumber` は付けない（`DEFAULT` レイアウトに番号 placeholder が増えるため）。
- placeholder には **座標（x y w h）と `name` `type` 以外を置かない**。PptxGenJS は placeholder 指定の `addText` で、レイアウト側の options を呼び出し側に**総取りで上書き**する（`Object.assign({}, itemOpts, placeHold.options)`、`pptxgen.cjs.js` L2537。`objectName` の上書きはその一例）。`fontFace` `lang` `align` `valign` `fontSize` `color` `margin` をここに置くと、run 側の実測値が静かに負ける。
- `defineSlideMaster` を呼ぶ順番が slideLayout の番号を決める。順番はこのファイルの配列の順で固定する。

### 5.4 マスター差し替えの継ぎ目（今回は no-op）

`replaceMaster` Patch 1 つに閉じる。入力は 3 つ。

```ts
interface MasterSwap {
  template: Buffer                                   // 既存 .pptx
  layoutMap: Record<string, string>                  // 'cover' → テンプレートの <p:cSld name>
  placeholderMap: Record<string, Record<'title' | 'body' | 'body2', { idx: number; type: string }>>
}
```

やること（調査 §2 の見立て通り）: テンプレートの master / layouts / theme / media を取り出して置く → `slideN.xml.rels` の `slideLayout` 関係（`Type` で引く）の Target を `layoutMap` で付け替える → スライドの `<p:ph idx type>` を `placeholderMap` で付け替える → `rebuildContentTypes`（後処理 7）が Override を作り直し、C10 が `r:id` の逆引きを検査する。
run に `fontFace` を付けている限り、差し替え後もフォントは游ゴシックのまま（§6）。テンプレートのフォントに寄せたければ、run の `<a:latin>/<a:ea>` を消す変換を同じ場所に足す。

- 得るもの: 差し替えに触る場所が 1 つの変換に閉じ、今の生成コードを変えずに済む。
- 失うもの: no-op の変換は受入テストに載らず、周りの変換が変わったときに前提が腐る。「no-op でも `MasterSwap` の型と、rels を `Type` で引く補助関数だけは実装して単体テストに載せる」を実装フェーズの要件にする。

### 5.5 スライドの背景

- 色: `SlideCapture.backgroundColor`（`.print-slide-container` の computed。透明（`rgba(…, 0)`）なら `FFFFFF`）。
- `frontmatter.background`（theme-default の `cover` `intro` が使う）は Node 側で `handleBackground` と同じ分岐をする: 値が `#` `rgb` `hsl` で始まれば**色**として `slide.background = { color }`（fetch も矩形も無し）。それ以外は**画像**として fetch し `slide.background = { data }`。
- 画像のとき、`handleBackground` は `linear-gradient(#0005, #0008), url(...)` で暗くする重ね（黒の不透明度 33〜53%）を入れ、文字を白にする。同じ見え方にするため、背景画像があるスライドには spTree の**最初の図形として**全面の矩形 `{ color: '000000', transparency: 55 }`（不透明度 45%）を置く。名前は `Background dim`（§4.5 の列の先頭に入れる）。
  - 得るもの: 白い文字が読める。
  - 失うもの: PowerPoint で背景を替えたい人は、矩形も併せて消す必要がある。名前で分かるようにしておく。また Slidev は `background-size: cover`（切り取り）だが、PptxGenJS の背景は `<a:stretch><a:fillRect/>`（引き伸ばし）なので、縦横比が 16:9 でない画像は歪む。歪みを避けたければ Node 側で 16:9 に切り出してから渡す（今回はやらない。`Report.dropped['background-crop']` に件数を出す）。
- `image-right` など、`.slidev-layout` の外の `background-image` は ppt-components.md §2.2 の規則で画像要素になる。`background-image` が `url()` を含まない（グラデーションだけ）なら画像への置き換え（`reason: 'gradient'`）。

---

## 6. フォント

**決定: run ごとに `fontFace: '游ゴシック'` を付ける（案 a）。加えて `pptx.theme = { headFontFace, bodyFontFace }` にも同じ名前を入れる。**

| | 案 a: run ごと | 案 b: テーマ任せ |
|---|---|---|
| 得るもの | どの環境でも `<a:latin>/<a:ea>/<a:cs>` に名前が残り、マスター差し替え後も変わらない。PptxGenJS の API だけで済む | 差し替えたテンプレートのフォントに自動で寄る。XML が小さい |
| 失うもの | 英数字も游ゴシックの Latin グリフになる。「英数字は別フォント」にしたければ後処理で `<a:latin>` を書き換える。差し替え時にテンプレートのフォントへ寄らない | PptxGenJS の theme は `<a:latin>` しか変えられず、和文名は theme1.xml の後処理が要る。表のセルなど run に既定の名前が入る箇所との不整合が出うる |

- 選ぶ理由: 差し替えは今回の範囲外で、まず「必ず游ゴシックになる」ことが要る。名前の付与は `fontFor(kind: 'body' | 'code')` の 1 関数に集め、案 b へ切り替えるときはこの関数が `undefined` を返すだけにする。
- テーマにも同じ名前を入れる理由: run に `fontFace` の付かない箇所（自動追加された placeholder、ノート）を拾うため。
  - 失うもの: 案 b に切り替えるとき、run とテーマの 2 か所を戻すことになる。`fontFor` と `themeFonts()` を同じ `fonts.ts` に置いて 1 ファイルで済ませる。
- フォント名は日本語名 `游ゴシック`（PptxGenJS のテーマの Jpan 指定と同じ）。ウェイトは Regular を指定し、太字は `bold` で `游ゴシック Bold` に解決させる。Mac の OS 同梱（`游ゴシック体`）とは名前が一致しないので、Mac では Microsoft 365 のクラウドフォント頼み（調査 §4）。
- 等幅（inline code、コードブロック）は `Consolas`。和文は PowerPoint の代替に任せる（`W-FONT` は出さない。PptxGenJS が `<a:ea>` を別名にできないため）。
- `lang`: スライドの `lang` 属性があればそれ、無ければ CLI の `--lang`（既定 `ja-JP`）。表示には効かず、校正とハイフネーションに効く。
- Slidev の表示も游ゴシックに揃える: headmatter の `fonts: { sans: '游ゴシック', mono: 'Consolas', local: ['游ゴシック', 'Consolas'] }`（`local` は `@slidev/types` にある。Google Fonts の取得を止める）。実装フェーズで `slides.md` に入れる。
  - 得るもの: 計測時と PowerPoint の行数が揃い、§3.3 の縮小率が当たる。
  - 失うもの: 計測に使うブラウザ（Mac / Linux CI）に游ゴシックが無いと別フォントで測る。枠の大きさはブロック幅で決まるのでほぼ影響しないが、行数（縮小率 §3.3）はずれうる。

---

## 7. CLI

パッケージ `packages/slidev-addon-pptx` に bin `slidev-pptx` を置く。

```
slidev-pptx export [entry=slides.md]
  --output, -o <path>     既定 ./slides-export.pptx（Slidev の既定名と同じ）
  --range <1-3,5>         Slidev の range と同じ綴り
  --theme <name>          resolveOptions に渡す
  --lang <tag>            run の lang の既定。既定 ja-JP
  --wait <ms>             描画後の追加待機。既定 0
  --timeout <ms>          ページ読み込みの上限。既定 30000
  --report <path>         書き出しの記録の JSON。既定 <output>.report.json
  --no-check              OPC 整合チェックを飛ばす
  --strict                警告が 1 つでもあれば exit 1
  --keep-server           失敗時にサーバとブラウザを閉じない（調査用。成功したときは閉じる）
slidev-pptx check <file.pptx>    OPC 整合チェックだけ
```

引数の検査は Slidev を立てる前に済ませる: `--wait` `--timeout` が数値でない、`--range` が `N` / `N-M` をコンマで並べた形でない（`N ≥ 1`、`M ≥ N`）なら exit 2。entry が無ければ exit 1（引数の綴りは合っているので 2 ではない）。

`package.json` の scripts:

```json
"export:pptx":       "slidev-pptx export slides.md",
"export:pptx-image": "slidev export --format pptx"
```

`export:pptx` の意味が「画像書き出し」から「ネイティブ書き出し」に変わる。画像書き出しは `export:pptx-image` に改名して残す。
- 得るもの: 「PPTX を出す」と言えば編集できるものが出る。親チケットの目的そのもの。
- 失うもの: 既存の手順書や CI が同じコマンドで別物を得る。画像書き出しに戻したい人は新しい名前を覚え直す。README（フェーズ 5）で改名を先頭に書く。
exit code: 0 成功 / 1 書き出し失敗または OPC error（`--strict` なら警告も）/ 2 引数の誤り。

---

## 8. 書き出しの記録

### 8.1 形

標準出力（人向け）と JSON（`--report`）に同じ内容。

```
slides-export.pptx: 12 slides, 34 native, 3 replaced, 3 warnings, dropped by rule: code-highlight ×3, blockquote-border ×1
  slide 7  replaced  mermaid        .mermaid (s7-e2 "Replaced 2") → image 1120×640
  slide 8  replaced  math           .katex-display (s8-e3 "Replaced 3") → image 480×96
  slide 8  warning   W-MATH-INLINE  inline math "E=mc^2" flattened to text (s8-e1 "Body")
  slide 9  replaced  unknown-element <button> (s9-e2 "Replaced 2") → image 180×48
  slide 5  warning   W-CSS          box-shadow dropped (s5-e4 "Text 4")
  slide 3  warning   W-LAYOUT       layout "section" not in table → blank
```

```ts
interface Report {
  output: string; generatedAt: string; slidev: string; pptxgenjs: string
  slides: number; native: number; replaced: number
  replacements: { slide: number; elementId: string; name: string; reason: ImageElement['reason']; selector: string; width: number; height: number }[]
  warnings: (Warning & { slide: number; name?: string })[]
  dropped: Record<string, number>          // 規則として捨てたもの・警告に上げないと決めたものの件数（下の一覧）
  zoom: Record<number, number>             // zoom を使ったスライド
  slideMap: Record<number, number>         // PPTX の中の番号（slideN.xml）→ 元のスライド番号。--range で絞ると 2 つがずれる
  check: CheckResult[]                     // §3.4。`src/types.ts` の { rule, part, message, level: 'error' | 'warn' }
}
```

`dropped` のキーと単位: `code-highlight`（色付きのコードブロック 1 つにつき 1。行ではない）、`blockquote-border`（引用 1 つにつき 1）、`transition`（デッキで 1）、`background-crop`（背景画像のあるスライド 1 枚につき 1。§5.5）、`console-noise`（Slidev 52.19 + floating-vue の既知の雑音 "Failed to patch FloatingVue" の件数。ブラウザの console.error から除いたもので、描画には影響しない。捨てた装飾ではなく「警告に上げない」と決めたもの）。

`--range` で絞ったときの番号: 後処理（図形名 `ctx.shapeNames`、縮小率 `ctx.autofit`）は PPTX の番号で引き、`warnings` / `replacements` の `slide` は元の番号のまま。人は `slideMap` で突き合わせる。標準出力では 1 行目の直後に `slide numbers (pptx → source): 1→2, 2→5` の 1 行を出す（ずれが無ければ出さない）。

- 得るもの: 何が画像になり、何を出せなかったかを人が確かめられる。CI では JSON を見て `replaced` の増加を検知できる。
- 失うもの: 記録の形が増えるたびに README と受入テストを直す。`Report` の型を 1 か所に置き、JSON はその型をそのまま出す。

### 8.2 警告コード

| code | 意味 |
|---|---|
| `W-CSS` | 再現できない装飾を捨てた（box-shadow、グラデーション、transform（2D 回転以外）、text-transform、非均一 border、filter） |
| `W-MATH-INLINE` | インライン数式を文字に平坦化した |
| `W-NESTED-PPT` | PPT 部品の中の PPT 部品を無視した |
| `W-PPT-CONTENT` | PptText の中の表・画像を無視した |
| `W-LI-BLOCK` | 箇条書きの中の表・画像・引用を無視した |
| `W-LAYOUT` | 対応表に無いレイアウトを `blank` にした。または、候補（`roleHint`）の placeholder がそのレイアウトに無く自由配置にした |
| `W-OVERFLOW` | 枠に収まらないので縮小率を書いた |
| `W-OFFSLIDE` | スライドの外に出た要素を内側に寄せた |
| `W-IMAGE` | 画像を取得できず灰色の矩形にした |
| `W-DARK` | ダークテーマ固定のデッキを測った色のまま出した |
| `W-HIDDEN` | 表示されていない要素を飛ばした（`display:none` 以外の理由: キャンバス外、opacity 0、寄せた結果 w か h が 0 以下） |
| `W-LINK` | リンクを外した。出すのは次の 2 つ: 相対パスや数字でないアンカー（開発サーバの URL に解決されて PPTX に残るため。ppt-components.md §3.4）／ `--range` の範囲外のスライドへの内部リンク（PptxGenJS の `hyperlink.slide` は PPTX の中の順番で、外さないと rels が実在しない `slideN.xml` を指して C4 で止まる。実測） |
| `W-SHAPE` | `PptShape` の図形を定義どおりに出せなかった。出すのは次の 3 つ: 図形の名前が定義（187 種）に無いので `rect` にした（調整値も捨てた）／ 調整値のうち、定義の avLst に無い名前・数でない値・32 bit 整数の範囲外の値を捨てた ／ 矢じりの値が PowerPoint に無いので `arrow` にした（図形と `type="line"` の線の両方。§4.3） |
| `W-INLINE` | 段落の中の `svg` `img` `table` `pre` `.katex-display` など、枠の単位でしか置き換えられないものを無視した（ppt-components.md §2.2 の 11） |
| `W-RENDER` | 描画の失敗。出すのは次の 3 つ: Slidev がスライドの描画に失敗している（収集器がエラー表示の要素を見つけた。そのスライドの番号）／ ブラウザの `console.error` か `pageerror`（slide 0）／ UnoCSS のクラスが効かないまま測った（§1.2。slide 0） |

「規則として捨てるもの」（コードの色付け、引用の左線、`transition`、背景画像の切り取り）は警告にしない。
- 得るもの: 毎回出る雑音を消し、対処できる警告だけが残る。
- 失うもの: 利用者は色付けや縦線が出なかったことに警告では気付けない。代わりに `Report.dropped` に件数を出し、標準出力の 1 行目にも載せる。

`transition` の扱い（決定）: **`dropped['transition']` だけ**に出し、警告 `W-TRANSITION` は出さない。設計の初版は表に `W-TRANSITION` を挙げつつ本文で「警告にしない」と書いていて矛盾し、実装 i0001-04 は両方出している。
- 決める理由: `transition` は Slidev のデッキの大半に付いていて対処のしようが無い（PowerPoint に出す手段が無い）。警告にすると `--strict` が `transition` のあるデッキで必ず赤になり、`--strict` を CI で使えない。「対処できる警告だけが残る」の方針に従う。
- 得るもの: `--strict` が使える。警告の一覧が対処できるものだけになる。
- 失うもの: 画面切り替えが消えたことは `dropped` の件数（標準出力の 1 行目）でしか分からない。
- 実装への影響: `W-TRANSITION` を出している 1 行（`convert.ts`）と、それを期待する検査（`tests/e2e/export.test.ts`、`tests/cli/export.test.ts` の `--strict` の検査は別の警告で組み直す）、README の「`--strict` は `transition` で必ず赤」の注意書きを直す。次の実装の子チケットで行う。

---

## 9. パッケージの置き場（実装フェーズの入力）

```
packages/slidev-addon-pptx/
  package.json            name: slidev-addon-pptx, bin: slidev-pptx, slidev.addon 用の設定
  components/             PptText.vue PptShape.vue PptImage.vue PptTable.vue（ppt-components.md）
  src/types.ts            Capture（§2）、Report（§8.1）、CheckResult（§3.4）の型。ブラウザ側と Node 側の両方が import する唯一の共有物
  src/collect/index.ts    ブラウザで動く収集器。Vue にも Node にも依存しない 1 ファイル。`export function collect(): Capture` を page.evaluate に渡す
  src/build/              masters.ts（`LAYOUTS` `defineMasters` `masterFor` = §5.1 の 2 段目）、layout.ts（`resolveLayout(slideIndex, data, layouts)` = §5.1 の 1 段目）、convert.ts（`build(capture, data): { pptx, ctx }`）、units.ts（§4.1。`emu` `pt` `inch` `clampBox` `toDashType` と、`textMargin()` = [l,r,b,t] / `cellMargin()` = [t,r,b,l] の 2 関数）、fonts.ts（`fontFor` `themeFonts`。§6）、names.ts（`assignNames`。§4.5）、sanitize.ts（`sanitizeXmlText`）、notes.ts（`notesText(note)`。§4.4）
  src/patch/              zip.ts（ZipView と、rels を `Type` で引く `findRelByType(relsDoc, type)`。§5.4）、patches/*.ts（§3.2 の 1 変換 1 ファイル、各ファイルが `Patch` を 1 つ export）、index.ts（`PATCHES: Patch[]` と `postProcess`）
  src/shapes/             presets.ts（PowerPoint の図形 187 種の定義。scripts/gen-presets.mjs の生成物で手で直さない。docs/adr/0002）、types.ts（定義の型）、geometry.ts（`PRESET_NAMES` `isPreset` `evalPreset`。定義を枠の大きさで評価して SVG の path と文字の枠を返す。DOM に依存しない。PptShape.vue と convert.ts が使う）、`normalizeAdjust`（調整値の正規化。定義に無いキー・数でない値・丸めた後に 32 bit 整数の範囲外の値を捨て、整数に丸める。PptShape.vue と convert.ts が同じ判定を使う）。後処理 4a（applyShapeAdjust.ts）が import するのは presets.ts だけ
  scripts/gen-presets.mjs presetShapeDefinitions.xml → src/shapes/presets.ts
  src/opc/check.ts        `export function check(buf: Buffer): Promise<CheckResult[]>`（§3.4）
  src/cli.ts              §7
  src/export.ts           `exportPptx(options): Promise<Report>`。§1 の流れ
```

依存: `pptxgenjs`、`jszip`（今は PptxGenJS の推移的依存。直接使うので直接依存にする）、`@xmldom/xmldom`（未インストール）、`playwright-chromium`（ルートにある）。
`slides.md` の headmatter に `addons: [slidev-addon-pptx]`、ルートの `package.json` に `"slidev-addon-pptx": "workspace:*"`、`pnpm-workspace.yaml` に `packages: [packages/*]`（今は `packages:` が無い）。
`package.json` `pnpm-lock.yaml` `pnpm-workspace.yaml` の変更は親チケットの `ask`。上の依存追加もその対象。

---

## 10. 制約との突き合わせ

| 調査の制約 | この設計での扱い |
|---|---|
| 後処理は jszip で開いて XML を触る一本道 | §3.1 |
| `cNvPr id` / `name` の重複を必ず直す。図形の特定は `objectName` で | 後処理 1、必須。**ただし `objectName` は placeholder 枠には効かない**（§3.2 の説明）。名前は後処理で付け直す |
| 未使用 placeholder が必ず入る | 後処理 2 |
| `[Content_Types].xml` の実在しない Override | 後処理 7 |
| 画像 placeholder は未実装 | 画像は全て座標指定 `addImage`（§4.3） |
| `write()` は無圧縮 | 後処理の再圧縮で DEFLATE |
| 公開 API は `resolveOptions` `createServer` `parser` のみ | §1.2。`go()` は自前で書く |
| export モードでしか `/print` が無い | §1.2 |
| レイアウト名の決定規則。**「DOM の class から読むほうが確実」** | §5.1。**この制約は調査の見立てが誤り。** `two-cols` の class は `two-columns`、`image-right` は `.slidev-layout.default` を内包するので、class から名前は決められない。Node 側で `getLayouts()` を使って解く。`wip/research/summary.md` §3 のこの 1 文の修正を親に依頼する（このチケットの `allow` は `wip/design/*` と `docs/adr/*` のみ） |
| クリック数は描画後 | 今回は `?print=true` 1 枚で不要 |
| 実測は `.print-slide-container` 基準、`zoom` の scale | §2、ppt-components.md §2.3（rect はそのまま、computed の長さに `zoom` を掛ける） |
| キャンバス高さ 551/552 | §4.1 |
| フォントの 2 案 | §6 |

親チケットの「見出しはタイトル枠」は、スライドの主見出し 1 つ（文書順で最初の h1 か h2）を指すと解した。2 つ目以降の見出しと h3 以降は本文枠の中の見出し段落になる（ppt-components.md §3.1）。

## 11. 要確認（実装・受入テストで最初に潰す）

実装 i0001-04 と受入テスト i0001-03 で出た答え（2026-09-13）:

1. `@xmldom/xmldom` の parse → serialize が PptxGenJS の XML を同値のまま戻すか（名前空間宣言、空要素、`xml:space`）。
   → **バイト列は同じにならない**（空要素が `<x/>` に畳まれ、属性の改行が詰まり、テキストの CRLF が LF になる）。DOM としては等価で、2 回目以降は冪等。「等価」の定義（要素名・属性の集合・テキストが再帰的に同じ。空白だけのテキストノード・コメント・XML 宣言は見ない）は `tests/patch/pipeline.test.ts` にある。§3.1 の「同値」はこの意味に読み替える。
2. コードブロックの行頭の空白が `<a:t>` で保たれるか。
   → **保たれる**（`"  return a"` が生成と e2e の両方で一致）。
3. 後処理 2 で消した placeholder のスライドを PowerPoint で開いたとき「修復」が出ないか。
   → **出ない**。利用者が Windows の PowerPoint で `slides.md` の書き出し 13 枚を開き、「修復」は出ず、編集もできた（2026-09-13）。
4. `getComputedStyle` を run ごとに呼ぶコスト。
   → 12 枚のデッキで `page.evaluate(collect)` は体感 1 秒未満。50 枚での別計測はしていない。
5. `locator.screenshot({ omitBackground: true })` が祖先の塗りまで透明にするか。
   → **未確認**。置き換え画像は撮れていて PowerPoint でも開ける（3 と同じ確認）が、`plain.md` も `slides.md` も置き換え要素が白背景の上にあるため、透けているかは見分けが付かない。色付きの箱の上に数式を置くデッキで確かめる。
6. `_slideObjects` の順 = `<p:spTree>` の順、自動追加 placeholder は末尾、が保たれるか。
   → 表・画像・図形が混ざるスライドでも**成立**（`tests/build/convert.test.ts` の「spTree の図形数」）。版を上げたら再確認。
7. `margin` の `[l, r, b, t]` が版を上げて直っていないか。
   → 4.0.1 では `[l, r, b, t]` のまま（`tests/fixtures/gen-pptx.test.ts` の「テキスト枠の margin は [l, r, b, t] の順」で固定。版を上げたらここが先に赤になる）。

答えが出たもの（この設計を書く過程の実測）: `fit: 'shrink'` + placeholder で `<a:normAutofit/>` は出る。表の `border` 配列で下線だけ出せる。

## 12. `CONTEXT.md` への追加を親に依頼する用語

- **収集結果**（Capture）: ブラウザで測った、スライドごとの要素の列。単位はキャンバス px
- **後処理**: PptxGenJS の出力 ZIP に対する、パス指定の XML 変換の列
- **OPC 整合チェッカ**: PPTX の ZIP・rels・Content_Types・id の整合を点検する検査
- **書き出しの記録**: 画像への置き換え一覧と警告を、標準出力と JSON に出したもの
- **流し込みブロック / 区切りブロック**: 1 つのテキスト枠にまとめて流す要素と、枠を閉じる要素（ppt-components.md §3.1）
