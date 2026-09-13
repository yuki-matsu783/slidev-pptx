# 受入テスト（i0001-03）

設計 `wip/design/native-export.md` と `wip/design/ppt-components.md` を検査として固定したもの。
実装（`packages/slidev-addon-pptx/`）はフェーズ 4 で入る。それまでは import 先が無いので全部 red。

## 回し方

依存（`vitest` `@xmldom/xmldom`、直接依存の `jszip` `pptxgenjs`）はフェーズ 4 の子チケットで入る。
入ったあと:

```sh
pnpm test        # = vitest run -c tests/vitest.config.ts      単体（ブラウザ無し）
pnpm test:e2e    # = vitest run -c tests/vitest.e2e.config.ts  e2e と CLI（playwright-chromium と Slidev サーバを使う）
```

`package.json` の scripts には `vitest run -c …` の形で書く（フェーズ 4）。`node` が PATH に無い環境があるので、呼ぶ側は常に `pnpm` 経由。

## 置き場

| 場所 | 何を固定するか | 設計の節 |
|---|---|---|
| `fixtures/gen-pptx.mjs` | PptxGenJS 4.0.1 だけで作る入力 PPTX。末尾で 4 条件を assert。`node tests/fixtures/gen-pptx.mjs` で再生成 | §3.2、§11 の 6・7 |
| `fixtures/pptx/sample.pptx` | 上の生成物（コミット済み。版を上げたら再生成） | |
| `fixtures/capture/basic.json` `.data.json` | 収集結果の見本と `options.data` 相当。対で使う | §2 |
| `fixtures/deck/plain.md` | 部品を使わない e2e 用デッキ（アドオン無しで起動する） | ppt-components §2.2 の 0〜15 |
| `fixtures/deck/components.md` | PPT 部品のデッキ。フェーズ 4 まで `describe.skip` | ppt-components §1 |
| `fixtures/deck/dark.md` | `colorSchema: dark`（`W-DARK`） | §1.2 |
| `opc/check.test.ts` | C1〜C14 を「通る / 落ちる」の対で | §3.4 |
| `patch/*.test.ts` | 後処理 1〜7 を 1 変換 1 ファイル。`pipeline.test.ts` は列の順、往復の同値、DEFLATE | §3.1、§3.2 |
| `build/*.test.ts` | 単位、名前、sanitize、マスター、フォント、ノート、レイアウト解決、Capture → XML | §4、§5、§6 |
| `e2e/collect.test.ts` | 収集器を `page.evaluate(collect)` で走らせる | ppt-components §2、§3 |
| `e2e/components.test.ts` | PPT 部品のデッキ。アドオンの実装が入るまで `describe.skip` | ppt-components §1 |
| `patch/replaceMaster.test-d.ts` | `MasterSwap` の型（`typecheck` で検査。実行時の `expectTypeOf` は no-op） | §5.4 |
| `e2e/export.test.ts` | 全経路。OPC 検査、Report、警告コード、dropped | §1、§8 |
| `cli/export.test.ts` | exit code、フラグ、scripts の綴り | §7 |

## import する名前（設計 §9 のとおり。検査で新しい名前を決めない）

`src/types.ts`: `Capture` `Report` `CheckResult` と要素の型
`src/collect/index.ts`: `collect`（自己完結。`page.evaluate` にそのまま渡す）
`src/build/`: `build`（convert.ts）、`LAYOUTS` `defineMasters` `masterFor`（masters.ts）、`resolveLayout`（layout.ts）、
`SLIDE_W_EMU` `SLIDE_W_IN` `emu` `pt` `inch` `clampBox` `toDashType` `textMargin` `cellMargin`（units.ts）、
`fontFor` `themeFonts`（fonts.ts）、`assignNames`（names.ts）、`sanitizeXmlText`（sanitize.ts）、`notesText`（notes.ts）
`src/patch/`: `Patch` `PatchContext` `PATCHES` `postProcess` `xmlPatch`（index.ts）、`findRelByType`（zip.ts）、
`patches/<name>.ts` が同名の `Patch` を 1 つ export、`replaceMaster.ts` は `MasterSwap` の型も
`src/opc/check.ts`: `check`
`src/export.ts`: `exportPptx({ entry, output, range?, lang?, check?, report? })`

検査が前提にしている引数の形と、設計に無い細部（実装が別解を選ぶなら検査を直す。設計に書き戻すべきものは design-feedback へ）:
- `build(capture, data, { assets: Record<captureId, dataUrl>, lang, layouts: string[] })` → `{ pptx, ctx }`。`layouts` は `options.utils.getLayouts()` のキー一覧
- `clampBox(box, canvas)` → `{ box, shift, dropped }`。`shift` は寄せた量と縮めた量の最大値
- `assignNames(elements, confirmedRoles: Map<id, 'title' | 'body' | 'body2'>, hasBackgroundDim)` → `string[]`。`Placeholder N` の N は spTree の 1 始まりの位置
- `resolveLayout(slideIndex, data, layoutNames: string[])` → レイアウト名
- `emu(px, canvas)` `pt(px, canvas)` `inch(px, canvas)` `textMargin(inset, canvas)` `cellMargin(inset, canvas)` `toDashType(dash)`
- `defineMasters` は `LAYOUTS` の写しを PptxGenJS に渡す（`createSlideMaster` が options を破壊するため）
- `sanitizeXmlText` は XML 1.0 で合法な DEL / C1（0x7F–0x9F）も落とす
- `notesText` は字下げ後の `- ` も `• ` にし、CRLF を LF にし、`undefined` を空文字にする
- `findRelByType` は `slideLayout` のような短い名前でも完全な URI でも引ける
- `two-cols` の `body2` placeholder の `type` は `body`
- CLI: entry が無ければ exit 1（引数の綴りは合っているので 2 ではない）
- `W-CSS` の message には捨てた property 名（`shadow` など）を含める
- `collect` は 1 関数の中に閉じる（Playwright は関数を文字列化して渡すので、モジュール先頭の定数やヘルパを参照できない）
- `Capture.canvas.height` は `.print-slide-container` の rect のまま（551〜552。丸め規則は設計に無い）
- `dropped['code-highlight']` の単位（ブロックか行か）は設計に無いので、検査は下限だけ見る
- e2e の viewport は Slidev と同じ「980 × 552 × 枚数」。待機列は `native-export.md` §1.2 を写し、UnoCSS の遅延注入を待つ段（`.slidev-layout` の padding が効く + `<style>` の長さが 500 ms 動かない）を足した（`tests/e2e/helpers.ts`）
- xmldom は XML の行末正規化で `<a:t>` の CRLF を LF にする。後処理 5 の区切りは `\r\n | \r | \n` のどれでも
- 往復の「等価」= 要素名・属性の集合・テキスト（行末を LF に揃える）が再帰的に同じ。空白だけのテキストノード・コメント・XML 宣言は見ない（`tests/patch/pipeline.test.ts` の `firstDifference`）
- e2e の待機は、デッキ由来のクラスが UnoCSS で生成されたことを番人にする（`plain.md` は `.pt-12` の `paddingTop: 48px`）。Slidev 自身の規則（`px-14`）はファイル変換時に展開されるので番人にならない。生成が来ない run は 30 秒で落ちる（黙って誤った値を測らない）
- CLI の検査は `process.execPath` + `packages/slidev-addon-pptx/bin/slidev-pptx.mjs` で起動する（Windows の `.CMD` は Node 22 では `shell: true` 無しに spawn できない）
- Slidev のサーバを立てる間は `NODE_ENV=development` にする（vitest の `test` や `production` だと、Vite 開発サーバで UnoCSS がデッキ由来のクラスを生成せず、`.pt-12` などが永遠に効かない。実測）。`tests/e2e/helpers.ts` と `src/export.ts` の両方
- vitest の typecheck は `tests/tsconfig.json` を使う（ルートの `tsconfig.json` は範囲外）。`typescript` と `@types/node` が要る（フェーズ 4 の依存に含める）
- `rebuildContentTypes` の拡張子 → ContentType は `jpg`/`jpeg` → `image/jpeg`、`webp` → `image/webp`（PptxGenJS 自身は `jpg` → `image/jpg` を出すが、IANA の型に揃える）
- `LAYOUTS` は Slidev の px を保持し、`defineMasters` が `canvasWidth` で換算する。placeholder の `idx` は PptxGenJS の採番で 100 始まり
- `notesText` は `[click]` の後ろの空白も詰める
- `data-ppt-box` で一部の辺だけ指定した部品も `boxSource: 'prop'`
- `transition` は `W-TRANSITION`（デッキで 1 回）と `dropped['transition']` の両方に出る（設計 §8.2 は表で警告に挙げつつ本文で「警告にしない」と書いていて矛盾している。design-feedback の候補）

## 設計へ書き戻す候補（design-feedback の子で出す）

- ppt-components.md §2.2: 「`p` の中身が `img`（と空白）だけなら、その `p` ごと画像要素（区切りブロック）」の特例。
  Markdown の `![]()` は `<p><img></p>` になり、今の規則では 11（`p`）が先に当たって 6（`img`）に届かない。
  このデッキは裸の `<img>` で回避している。`<a><img></a>` も 1 行に書くと段落に落ちるので 3 行に割ってある
- ppt-components.md §2.2 の 12: タグ一覧に `a` を足す。§3.3 の「`a > img` なら `link` を要素に」へ到達する行が無い
- ppt-components.md §2.2 の 1: `hr` は高さ 0〜1 px なので、「rect が空」を `w === 0 || h === 0` で判定すると規則 10 に届かない
- ppt-components.md §3.4: Slidev は `[x](#3)` を `href="##3"` にする。「`#N` か `/N`」に `##N` を足す。このデッキは `<a href="/3">` で回避している
- ppt-components.md §3.5: 罫線は `td`/`th` の computed では 4 辺とも 0（Slidev は `tr { border-b }`）。セルが 0 なら `tr` → `table` へ遡る。`tr` の色は `rgba(…, 0.2)` なので不透明度の扱いも決める
- ppt-components.md §3.6: 行強調 `{2}` で強調されない行は `.slidev-code-dishonored`（opacity 0.3）。§3.4 の「祖先の opacity の積 → transparency」を当てないよう、コード枠では opacity を無視する
- ppt-components.md §3.4: `transparency = 100 − opacity × 100` の丸め（整数に）
- native-export.md §3.2 の 5: 「`<a:t>` の CRLF」は DOM では LF。区切りの定義を `\r\n | \r | \n` に
- native-export.md §1.2: 待機列に UnoCSS の遅延注入を待つ段を足す
- native-export.md §8.2: `W-TRANSITION` を警告に出すか `dropped` だけにするかを 1 つに
- native-export.md §8.2: 警告コード `W-LINK` を足す。PptxGenJS の `hyperlink.slide` は PPTX の中での順番なので、`--range` で絞ると元の番号とずれる。範囲内なら写像し、範囲外へのリンクは外して `W-LINK`（外さないと rels が実在しない slideN.xml を指して C4 で落ちる。実測）
- native-export.md §7: `--range` は Slidev の `/print` では効かないことがある（`useNav` が初期化時に `query.range` を 1 度読むだけ）。URL に渡したうえで Node 側でも絞る（実測）

## 設計 native-export.md §11「要確認」への答え（フェーズ 4 の実測）

1. xmldom の往復: バイト列は同じにならない（空要素が `<x/>` に畳まれ、属性の改行が詰まり、テキストの CRLF が LF になる）。DOM としては等価で、2 回目以降は冪等（`tests/patch/pipeline.test.ts`）
2. コードの行頭の空白: `<a:t>` に保たれる（`"  return a"` が convert と e2e の両方で一致）
3. 空 placeholder を消したスライドの PowerPoint 実機: **未確認**（Windows が要る。人のレビューで）
4. `getComputedStyle` のコスト: 12 枚のデッキで `page.evaluate(collect)` は体感 1 秒未満。別計測はしていない
5. `omitBackground`: 置き換え画像は撮れているが、祖先の塗りが透けるかは**未確認**（plain.md の置き換え要素は白背景の上にあるため見分けが付かない）
6. spTree の順 = add した順、自動追加 placeholder は末尾: 表・画像・図形が混ざるスライドでも成立（`tests/build/convert.test.ts` の「spTree の図形数」）
7. `margin` の並び: 4.0.1 でも `[l, r, b, t]`（`tests/fixtures/gen-pptx.test.ts`）

## フェーズ 4 で最初に確かめること

- vitest が `tests/**` を ESM として動かすとき `import.meta.url` が取れること（`__dirname` は使わず `here` に統一。vitest の config も同じ）
- `zoom` は CSS の `scale` プロパティ（`getComputedStyle(el).scale === "0.8"`、`transform` は `none`）。`transform` を読む実装では取れない
- スライド 12 のような `<!-- -->` は Slidev がノートとして拾う。ノートの検査を足すときはデッキ側のコメントに注意
- `page.evaluate(collect)` の直列化（上の「1 関数の中に閉じる」）
- `plain.md` の枚数は `tests/e2e/slides.ts` の `SLIDE_COUNT`（12）。`@slidev/parser` は `---` で始まる行を無条件に区切りにするので、水平線は `***` で書いてある
