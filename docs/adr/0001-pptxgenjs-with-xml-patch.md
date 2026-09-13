---
status: accepted
date: 2026-09-13
ticket: i0001-02
updated: 2026-09-13 (i0001-05。実装 i0001-04 で確かめた制約を Consequences に追記)
---

# PptxGenJS で組み、足りない分は XML を後から足す。書き出しは一方向

Slidev の標準書き出しは各スライドを画像として貼るため、受け取った人が PowerPoint で中身を直せない。
ネイティブ書き出しの生成器として PptxGenJS を使い、PptxGenJS が出せないもの・壊すものは
書き出した ZIP の XML を後処理で直す。向きは `slides.md` → PPTX の一方向で、PowerPoint での
手直しは Slidev に戻さない。

理由: PptxGenJS はテキスト・図形・画像・表・テキストのプレースホルダー・ノート・ハイパーリンクまでを
API で出せて（画像や表のプレースホルダーは未実装）、OOXML のパートと rels の骨格を自分で書かずに済む。一方で `cNvPr id` の重複、
未使用プレースホルダーの自動追加、実在しないパートへの `[Content_Types].xml` の Override、
グループ化・アニメーション・画面切り替えの API 欠如があり（調査 `wip/research/summary.md` §1）、
これらは ZIP を開いて XML を直す以外に手が無い。後処理を「パス指定の XML 変換の列」として
1 層に切れば、PptxGenJS の版に依存する部分と OOXML 直書きの部分が分かれ、将来のマスター
差し替えも同じ層の 1 変換になる。

## Considered Options

| 案 | 退けた理由 |
|---|---|
| **XML を全部自前で書く**（PptxGenJS を使わない） | 骨格（presentation / master / layout / theme / notesMaster / rels / Content_Types）を全て自分で持つことになり、PowerPoint の「修復」を避けるための知識を最初から全部抱える。得るのは後処理の層が要らなくなることだけで、PptxGenJS が出せる範囲（テキスト・表・画像）でも同じ量の XML を書く。今の範囲では割に合わない |
| **PptxGenJS だけで済ませる**（後処理なし） | 表を 1 つ使うと `cNvPr id` が重複し、`masterName` 付きのスライドには空のプレースホルダーが必ず入る。どちらも API では避けられない。「修復」が出る PPTX は受け取った人の信頼を失うので、後処理なしは成り立たない |
| **双方向**（PowerPoint の編集を `slides.md` に戻す） | PPTX → Markdown の写像は一意でなく（図形の位置・書式の大半は Markdown に置き場が無い）、往復のたびに情報が落ちる。原本を 1 つに決めるほうが運用が単純。PowerPoint 側は「最後の手直し」と割り切る |

## Consequences

- PptxGenJS の API には「出せるように見えて壊れる」経路が 8 つある（要素リンクが rels に載らない、placeholder 指定で options が総取りされる、`margin` の並び、`rectRadius` の単位、SVG、負の数値、ノートの CRLF、run ごとの `<a:pPr>`）。
  いずれも後処理か呼び分けで押さえ、一覧と根拠は `wip/design/native-export.md` §4 と `tests/README.md` に「実測」として置く。版を上げるたびに受入テスト（`tests/fixtures/gen-pptx.mjs` の 4 条件）で崩れていないかを確かめる（実装 i0001-04 で追記）。
- 収集は Vite の開発サーバで描画した DOM を測る。UnoCSS はデッキ由来のクラスを最初の読み込みで生成しないことがあるので、書き出しは「クラスが効いたこと」を確認してから測り、効いていなければ再読み込みする。`NODE_ENV` が `test` / `production` だと生成されない（実装 i0001-04 で追記）。
- 後処理が XML を直接触るので、読めるだけでは足りず、OPC 整合チェッカ（Override・rels・id の突き合わせ）を自前で持って CI で回す（調査 `wip/research/summary.md` §1: 「修復」の原因は python-pptx では検出できなかった）。
- PptxGenJS の版を上げるときは、後処理が前提にしている出力の形（id の採番、Content_Types の書き方、ノートの CRLF）が変わっていないかを受入テストで確かめる。
- 一方向なので、PowerPoint で直した内容を残したい人は、Slidev 側を直して書き出し直す。設計文書 `wip/design/native-export.md` と `ppt-components.md` はこの決定に基づく。
