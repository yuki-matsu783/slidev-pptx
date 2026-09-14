// Capture → PptxGenJS の変換（native-export.md §4、§5.1、§5.5、§6、§8）。
// 見本 tests/fixtures/capture/basic.json + basic.data.json を build に通し、出力 XML を検査する。
import { beforeAll, describe, expect, it } from 'vitest'
import { build } from '../../packages/slidev-addon-pptx/src/build/convert'
import type { PatchContext } from '../../packages/slidev-addon-pptx/src/patch/index'
import { PATCHES, postProcess } from '../../packages/slidev-addon-pptx/src/patch/index'
import type { Capture, DeckData, LineElement, ShapeElement, SlideCapture } from '../../packages/slidev-addon-pptx/src/types'
import { emu, inch, pt } from '../../packages/slidev-addon-pptx/src/build/units'
import { openPptx, els, shapesOf, shapeNamed, cNvPrOf, textOf } from '../helpers/pptx'
import type { OpenedPptx } from '../helpers/pptx'
import captureJson from '../fixtures/capture/basic.json'
import dataJson from '../fixtures/capture/basic.data.json'

const canvas = { width: 980, height: 552 }
// build が capture / data を破壊的に触っても他の it に波及しないよう、毎回 clone を渡す
const data = () => structuredClone(dataJson)
const PNG_1x1 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

let raw: Buffer
let p: OpenedPptx
let ctx: PatchContext
let slide: Record<number, Document> = {}
let rels: Record<number, Document> = {}

beforeAll(async () => {
  // build が Capture を破壊的に触っても他の it に波及しないよう、毎回 clone を渡す
  const capture = structuredClone(captureJson) as unknown as Capture
  const r = await build(capture, data(), { assets: { 's2-e9': PNG_1x1 }, lang: 'ja-JP', layouts: dataJson.layouts })
  ctx = r.ctx
  raw = await r.pptx.write({ outputType: 'nodebuffer' }) as Buffer
  p = await openPptx(raw)
  for (let i = 1; i <= 5; i++) {
    slide[i] = await p.xml(`ppt/slides/slide${i}.xml`)
    rels[i] = await p.xml(`ppt/slides/_rels/slide${i}.xml.rels`)
  }
})

const ph = (doc: Document, type: string) => shapesOf(doc).find((s) => els(s, 'p', 'ph')[0]?.getAttribute('type') === type)
const phIdx = (doc: Document, idx: string) => shapesOf(doc).find((s) => els(s, 'p', 'ph')[0]?.getAttribute('idx') === idx)
const relTarget = (doc: Document, id: string) => els(doc, 'rel', 'Relationship').find((r) => r.getAttribute('Id') === id)?.getAttribute('Target')
const rPrs = (el: Element) => els(el, 'a', 'rPr')
const paras = (el: Element) => els(el, 'a', 'p')

describe('convert: レイアウトと placeholder（§5.1、§4.2）', () => {
  it('cover: title と body の候補が placeholder に入る（見出しは div.my-auto 越しでも候補になる）', () => {
    const t = ph(slide[1], 'title')!
    expect(textOf(t)).toBe('Slidev 入門')
    expect(textOf(ph(slide[1], 'body')!)).toBe('Markdown で書くスライド')
  })
  it('placeholder 枠には座標を渡さない（レイアウトの値が勝つ）: cover の title は y ≈ 2.99 in', () => {
    const off = els(ph(slide[1], 'title')!, 'a', 'off')[0]
    expect(Math.abs(Number(off.getAttribute('y')) - 2.99 * 914400)).toBeLessThanOrEqual(0.01 * 914400)
  })
  it('two-cols: body2 の候補は idx=102 の placeholder に入る', () => {
    expect(textOf(phIdx(slide[3], '102')!)).toBe('長い本文')
  })
  it('対応表に無い section は blank で、見出しは自由配置（placeholder 無し）+ W-LAYOUT', () => {
    expect(shapesOf(slide[4]).some((s) => els(s, 'p', 'ph').length)).toBe(false)
    expect(relTarget(rels[4], relIdOfLayout(rels[4]))).toBe('../slideLayouts/slideLayout2.xml')
    expect(ctx.report.warnings.some((w) => w.code === 'W-LAYOUT' && w.slide === 4)).toBe(true)
  })
  it('getLayouts() に無い綴り違いは default として扱い、警告は出ない', () => {
    expect(relTarget(rels[5], relIdOfLayout(rels[5]))).toBe('../slideLayouts/slideLayout4.xml')
    expect(textOf(ph(slide[5], 'title')!)).toBe('綴り違い')
    expect(ctx.report.warnings.some((w) => w.slide === 5)).toBe(false)
  })
  it('自由配置の枠は x y w h を EMU で持ち、txBox="1"', () => {
    const code = shapesOf(slide[2]).find((s) => textOf(s).startsWith('const a'))!
    const off = els(code, 'a', 'off')[0]
    expect(Number(off.getAttribute('x'))).toBe(emu(56, canvas))
    expect(Number(off.getAttribute('y'))).toBe(emu(230, canvas))
    expect(els(code, 'p', 'cNvSpPr')[0].getAttribute('txBox')).toBe('1')
    expect(els(ph(slide[2], 'title')!, 'p', 'cNvSpPr')[0].getAttribute('txBox')).not.toBe('1')
  })
})

describe('convert: 段落と run（§4.2）', () => {
  it('箇条書き: level 0 は lvl 無しで buChar ▪、level 1 は lvl="1"', () => {
    const body = ph(slide[2], 'body')!
    const ps = paras(body)
    const pPr0 = els(ps[0], 'a', 'pPr')[0]
    expect(pPr0.hasAttribute('lvl')).toBe(false)
    expect(els(pPr0, 'a', 'buChar')[0].getAttribute('char')).toBe('▪')
    expect(els(ps[1], 'a', 'pPr')[0].getAttribute('lvl')).toBe('1')
  })
  it('番号つき: buAutoNum に startAt、種類は既定の arabicPeriod', () => {
    const an = els(paras(ph(slide[2], 'body')!)[2], 'a', 'buAutoNum')[0]
    expect(an.getAttribute('startAt')).toBe('3')
    expect(an.getAttribute('type')).toBe('arabicPeriod')
  })
  it('run の書式: b / i / code は Consolas + highlight', () => {
    const body = ph(slide[2], 'body')!
    const all = rPrs(body)
    expect(all[0].getAttribute('b')).toBe('1')
    const italic = all.find((r) => r.getAttribute('i') === '1')!
    expect(italic).toBeTruthy()
    const code = all.find((r) => els(r, 'a', 'latin')[0]?.getAttribute('typeface') === 'Consolas')!
    expect(els(code, 'a', 'highlight')).toHaveLength(1)
  })
  it('ハイパーリンク: URL は External の rels、slide は slideN.xml を指す', () => {
    const body = ph(slide[2], 'body')!
    const links = els(body, 'a', 'hlinkClick')
    const targets = links.map((l) => relTarget(rels[2], l.getAttribute('r:id') ?? l.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')!))
    expect(targets).toContain('https://sli.dev')
    expect(targets).toContain('slide3.xml')
  })
  it('run の transparency は alpha に（50 → 50000）', () => {
    const body = ph(slide[1], 'body')!
    expect(els(body, 'a', 'alpha')[0].getAttribute('val')).toBe('50000')
  })
  it('font size は pt(px) の小数 1 桁 × 100', () => {
    const t = rPrs(ph(slide[2], 'title')!)[0]
    expect(t.getAttribute('sz')).toBe(String(Math.round(pt(30, canvas) * 100)))
  })
  it('lineSpacingMultiple は lineHeight / size（1.0 未満は 1.0）', () => {
    const body = ph(slide[2], 'body')!
    const pct = els(paras(body)[0], 'a', 'spcPct')[0]
    expect(Number(pct.getAttribute('val'))).toBe(Math.round((31.68 / 17.6) * 100000))
  })
  it('全枠に normAutofit', () => {
    for (const s of shapesOf(slide[2]).filter((s) => els(s, 'p', 'txBody').length && !els(s, 'a', 'tbl').length)) {
      expect(els(s, 'a', 'normAutofit').length, cNvPrOf(s).getAttribute('name')!).toBe(1)
    }
  })
})

describe('convert: フォントと lang（§6）', () => {
  it('本文の run には 游ゴシック（latin / ea / cs）', () => {
    const r = rPrs(ph(slide[2], 'title')!)[0]
    expect(els(r, 'a', 'latin')[0].getAttribute('typeface')).toBe('游ゴシック')
    expect(els(r, 'a', 'ea')[0].getAttribute('typeface')).toBe('游ゴシック')
  })
  it('lang はスライドの lang 属性、無ければ --lang の既定', () => {
    expect(rPrs(ph(slide[1], 'title')!)[0].getAttribute('lang')).toBe('ja')
    expect(rPrs(ph(slide[2], 'title')!)[0].getAttribute('lang')).toBe('ja-JP')
  })
  it('テーマの major / minor も 游ゴシック', async () => {
    const theme = await p.xml('ppt/theme/theme1.xml')
    expect(els(els(theme, 'a', 'majorFont')[0], 'a', 'latin')[0].getAttribute('typeface')).toBe('游ゴシック')
    expect(els(els(theme, 'a', 'minorFont')[0], 'a', 'latin')[0].getAttribute('typeface')).toBe('游ゴシック')
  })
})

describe('convert: 自由配置のテキスト枠の見た目（§4.2）', () => {
  it('コード枠: fill、roundRect と rectRadius（インチ）、margin は [l, r, b, t]', () => {
    const code = shapesOf(slide[2]).find((s) => textOf(s).startsWith('const a'))!
    expect(els(code, 'a', 'srgbClr').some((c) => c.getAttribute('val') === 'F8F8F8')).toBe(true)
    expect(els(code, 'a', 'prstGeom')[0].getAttribute('prst')).toBe('roundRect')
    const cx = Number(els(code, 'a', 'ext')[0].getAttribute('cx'))
    const cy = Number(els(code, 'a', 'ext')[0].getAttribute('cy'))
    const adj = els(code, 'a', 'gd')[0].getAttribute('fmla')
    expect(adj).toBe(`val ${Math.round((inch(6, canvas) * 914400 * 100000) / Math.min(cx, cy))}`)
    const bodyPr = els(code, 'a', 'bodyPr')[0]
    // inset [12, 16, 12, 16]（上右下左）→ 左右 16 px、上下 12 px
    expect(Number(bodyPr.getAttribute('lIns'))).toBe(Math.round(pt(16, canvas) * 12700))
    expect(Number(bodyPr.getAttribute('tIns'))).toBe(Math.round(pt(12, canvas) * 12700))
    expect(Number(bodyPr.getAttribute('lIns'))).toBeGreaterThan(Number(bodyPr.getAttribute('tIns')))
  })
  it('コードは 1 行 1 段落で Consolas、行頭の空白が残る', () => {
    const code = shapesOf(slide[2]).find((s) => textOf(s).startsWith('const a'))!
    const ps = paras(code)
    expect(ps).toHaveLength(2)
    expect(els(ps[1], 'a', 't')[0].textContent).toBe('  return a')
    expect(els(ps[1], 'a', 'latin')[0].getAttribute('typeface')).toBe('Consolas')
  })
})

describe('convert: 表（§4.3）', () => {
  it('gridCol は colW を EMU で、罫線は 4 辺明示（下だけ solid、他は noFill）', () => {
    const tbl = els(slide[2], 'a', 'tbl')[0]
    const cols = els(tbl, 'a', 'gridCol').map((c) => Number(c.getAttribute('w')))
    expect(cols).toEqual([emu(434, canvas), emu(434, canvas)])
    const cell = els(tbl, 'a', 'tc')[0]
    expect(els(els(cell, 'a', 'lnB')[0], 'a', 'solidFill')).toHaveLength(1)
    expect(els(els(cell, 'a', 'lnT')[0], 'a', 'noFill')).toHaveLength(1)
    expect(els(els(cell, 'a', 'lnL')[0], 'a', 'noFill')).toHaveLength(1)
  })
  it('セルの margin は上右下左 pt。上 0 px は 1 pt に切り上げる', () => {
    const tcPrs = els(els(slide[2], 'a', 'tbl')[0], 'a', 'tcPr')
    expect(Number(tcPrs[0].getAttribute('marT'))).toBe(Math.round(pt(12, canvas) * 12700))
    expect(Number(tcPrs[2].getAttribute('marT'))).toBe(12700)
  })
  it('セルの文字と align', () => {
    const tbl = els(slide[2], 'a', 'tbl')[0]
    const tcs = els(tbl, 'a', 'tc')
    expect(textOf(tcs[3])).toBe('通常')
    expect(els(tcs[3], 'a', 'pPr')[0].getAttribute('algn')).toBe('ctr')
  })
})

describe('convert: 画像・図形・線（§4.3）', () => {
  it('画像: descr に alt、要素リンクは cNvPr の hlinkClick で rels に登録される', () => {
    const pic = els(slide[2], 'p', 'pic').find((x) => cNvPrOf(x).getAttribute('descr') === '写真')!
    const link = els(cNvPrOf(pic), 'a', 'hlinkClick')[0]
    const rid = link.getAttribute('r:id') ?? link.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')!
    expect(relTarget(rels[2], rid)).toBe('https://example.com/photo')
  })
  it("fit: 'cover' は sizing（srcRect が出る）、'fill' は sizing 無し（stretch/fillRect だけ。write が落ちないこと）", () => {
    const cover = els(slide[2], 'p', 'pic').find((x) => cNvPrOf(x).getAttribute('descr') === '写真')!
    expect(els(cover, 'a', 'srcRect')).toHaveLength(1)
    const fill = els(slide[2], 'p', 'pic').find((x) => cNvPrOf(x).getAttribute('descr') === '引き伸ばし')!
    expect(els(fill, 'a', 'srcRect')).toHaveLength(0)
    expect(els(fill, 'a', 'fillRect')).toHaveLength(1)
  })
  it('文字のある図形: prstGeom、文字、要素リンクは run に写り、cNvPr には無い、dash dot → sysDot', () => {
    const arrow = shapesOf(slide[2]).find((s) => cNvPrOf(s).getAttribute('name') === 'arrow')!
    expect(els(arrow, 'a', 'prstGeom')[0].getAttribute('prst')).toBe('rightArrow')
    expect(textOf(arrow)).toBe('次へ')
    expect(els(cNvPrOf(arrow), 'a', 'hlinkClick')).toHaveLength(0)
    expect(els(rPrs(arrow)[0], 'a', 'hlinkClick')).toHaveLength(1)
    expect(els(arrow, 'a', 'prstDash')[0].getAttribute('val')).toBe('sysDot')
  })
  it('文字の無い図形: addShape で要素リンク（slide）が cNvPr に付き、fill の transparency が alpha に', () => {
    const rr = shapesOf(slide[2]).find((s) => els(s, 'a', 'prstGeom')[0]?.getAttribute('prst') === 'roundRect' && textOf(s) === '')!
    const link = els(cNvPrOf(rr), 'a', 'hlinkClick')[0]
    const rid = link.getAttribute('r:id') ?? link.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')!
    expect(relTarget(rels[2], rid)).toBe('slide1.xml')
    expect(els(rr, 'a', 'alpha')[0].getAttribute('val')).toBe('80000')
  })
  it('線: prst="line"、水平なので flipV 無し', () => {
    const line = shapesOf(slide[2]).find((s) => els(s, 'a', 'prstGeom')[0]?.getAttribute('prst') === 'line')!
    expect(els(line, 'a', 'xfrm')[0].hasAttribute('flipV')).toBe(false)
    expect(els(line, 'a', 'srgbClr').some((c) => c.getAttribute('val') === 'CCCCCC')).toBe(true)
  })
  it('置き換え画像: assets の PNG が pic として入る（写真 + 引き伸ばし + 置き換え = 3）', () => {
    const pics = els(slide[2], 'p', 'pic')
    expect(pics).toHaveLength(3)
  })
  it('assets に無い置き換えは灰色の矩形 + W-IMAGE', async () => {
    const capture = structuredClone(captureJson) as unknown as Capture
    const r = await build(capture, data(), { assets: {}, lang: 'ja-JP', layouts: dataJson.layouts })
    expect(r.ctx.report.warnings.some((w) => w.code === 'W-IMAGE' && w.slide === 2)).toBe(true)
    const q = await openPptx(await r.pptx.write({ outputType: 'nodebuffer' }) as Buffer)
    expect(els(await q.xml('ppt/slides/slide2.xml'), 'p', 'pic')).toHaveLength(2)
  })
})

describe('convert: 図形の種類・調整値・反転・矢じり（slide 4）', () => {
  const sp = (name: string) => shapesOf(slide[4]).find((s) => cNvPrOf(s).getAttribute('name') === name)!
  const xfrm = (name: string) => els(sp(name), 'a', 'xfrm')[0]
  const shapeWarnings = (id: string) => ctx.report.warnings.filter((w) => w.code === 'W-SHAPE' && w.slide === 4 && w.elementId === id)

  it('shapeNames: 自由配置の見出しのあとに図形 6 つ', () => {
    expect(ctx.shapeNames[4]).toEqual(['Text 1', 'Shape 2', 'Shape 3', 'Shape 4', 'Shape 5', 'Shape 6', 'Shape 7'])
  })
  it('PptxGenJS の ShapeType に無いコネクタ（bentConnector3）と foldedCorner も prst 名のまま出る', () => {
    expect(els(sp('Shape 2'), 'a', 'prstGeom')[0].getAttribute('prst')).toBe('bentConnector3')
    expect(els(sp('Shape 3'), 'a', 'prstGeom')[0].getAttribute('prst')).toBe('foldedCorner')
    expect(textOf(sp('Shape 3'))).toBe('メモ')
  })
  it('未知の名前は rect + W-SHAPE 1 件（調整値も捨てたと書く）', () => {
    expect(els(sp('Shape 4'), 'a', 'prstGeom')[0].getAttribute('prst')).toBe('rect')
    const ws = shapeWarnings('s4-e4')
    expect(ws.length).toBe(1)
    expect(ws[0].message).toContain('調整値も捨てた')
    expect(ctx.adjust[4]['Shape 4']).toBeUndefined()
  })
  it('flipH / flipV は xfrm に（文字あり = addText、文字なし = addShape の両方）', () => {
    expect(xfrm('Shape 3').getAttribute('flipH')).toBe('1')
    expect(xfrm('Shape 3').getAttribute('flipV')).toBe('1')
    expect(xfrm('Shape 5').hasAttribute('flipH')).toBe(false)
    expect(xfrm('Shape 5').getAttribute('flipV')).toBe('1')
    expect(xfrm('Shape 2').hasAttribute('flipH')).toBe(false)
  })
  it('arrow の head / tail は headEnd / tailEnd に（知らない値は arrow にして W-SHAPE）', () => {
    const ln = els(sp('Shape 2'), 'a', 'ln')[0]
    expect(els(ln, 'a', 'headEnd')[0].getAttribute('type')).toBe('triangle')
    expect(els(ln, 'a', 'tailEnd')[0].getAttribute('type')).toBe('arrow')
    expect(shapeWarnings('s4-e2').filter((w) => w.message.includes('bogus'))).toHaveLength(1)
    expect(shapeWarnings('s4-e3').some((w) => w.message.includes('矢じり'))).toBe(false)
  })
  it('adj は整数に丸めて ctx.adjust に入り、定義の avLst に無いキーは捨てて W-SHAPE', () => {
    expect(ctx.adjust[4]['Shape 2']).toEqual({ adj1: 25000 })
    const w = shapeWarnings('s4-e2').find((x) => x.message.includes('foo'))!
    expect(w).toMatchObject({ name: 'Shape 2' })
  })
  it('32 bit 整数の範囲外の adj は捨てて W-SHAPE', () => {
    expect(ctx.adjust[4]['Shape 3']).toBeUndefined()
    expect(shapeWarnings('s4-e3')).toHaveLength(1)
  })
  it('roundRect で adj.adj があるときは radius より adj が勝ち、rectRadius の gd は出ない（後処理が書く）', () => {
    expect(els(sp('Shape 5'), 'a', 'prstGeom')[0].getAttribute('prst')).toBe('roundRect')
    expect(els(sp('Shape 5'), 'a', 'gd')).toHaveLength(0)
    expect(ctx.adjust[4]['Shape 5']).toEqual({ adj: 30000 })
  })
  it('roundRect の radius（px）は rectRadius を使わず、寄せた後の枠の短辺で adj に換算する', () => {
    // slide 2: 100×40、radius 8 → 8 / 40
    expect(ctx.adjust[2]).toEqual({ 'Shape 7': { adj: 20000 } })
    const rr = shapesOf(slide[2]).find((s) => els(s, 'a', 'prstGeom')[0]?.getAttribute('prst') === 'roundRect' && textOf(s) === '')!
    expect(els(rr, 'a', 'gd')).toHaveLength(0)
    // radius 0 は 0（PptxGenJS の rectRadius は 0 を捨てて PowerPoint の既定の角丸になる）
    expect(ctx.adjust[4]['Shape 6']).toEqual({ adj: 0 })
    expect(els(sp('Shape 6'), 'a', 'gd')).toHaveLength(0)
    // x=-40 w=100 h=80 → 寄せて 60×80、radius 10 → 10 / 60
    expect(ctx.adjust[4]['Shape 7']).toEqual({ adj: 16667 })
  })
  it('後処理のあと、avLst の gd は ctx.adjust の値になる', async () => {
    const q = await openPptx(await postProcess(raw, PATCHES, ctx))
    const s2 = await q.xml('ppt/slides/slide2.xml')
    const s4 = await q.xml('ppt/slides/slide4.xml')
    expect(gdsOf(s2, 'Shape 7')).toEqual([['adj', 'val 20000']])
    expect(gdsOf(s4, 'Shape 2')).toEqual([['adj1', 'val 25000']])
    expect(gdsOf(s4, 'Shape 5')).toEqual([['adj', 'val 30000']])
    expect(gdsOf(s4, 'Shape 6')).toEqual([['adj', 'val 0']])
    expect(gdsOf(s4, 'Shape 7')).toEqual([['adj', 'val 16667']])
  })
})

describe('convert: 合成の Capture（範囲指定・調整値の範囲）', () => {
  const shape = (id: string, o: Partial<ShapeElement> = {}): ShapeElement => ({
    id, name: '', source: 'ppt', kind: 'shape', shape: 'roundRect',
    box: { x: 10, y: 10, w: 100, h: 40 }, boxSource: 'prop', frame: { fill: { color: '#dddddd' }, inset: [0, 0, 0, 0] }, ...o,
  })
  const slideOf = (no: number, elements: SlideCapture['elements']): SlideCapture => ({ no, lang: 'ja-JP', zoom: 1, backgroundColor: '#ffffff', elements, warnings: [] })
  const deck = (n: number): DeckData => ({ slides: Array.from({ length: n }, (_, i) => ({ index: i, frontmatter: {} })), layouts: ['default'] })
  const run = async (slides: SlideCapture[], n: number) => {
    const r = await build({ canvas, slides }, deck(n), { assets: {}, lang: 'ja-JP', layouts: ['default'] })
    return { ctx: r.ctx, out: await postProcess((await r.pptx.write({ outputType: 'nodebuffer' })) as Buffer, PATCHES, r.ctx) }
  }

  it('スライド番号が連番でない（範囲指定の 3 と 5）とき、ctx.adjust は PPTX の番号で引け、slide2.xml に gd が入る', async () => {
    const { ctx: c, out } = await run([slideOf(3, [shape('s3-e1', { adj: { adj: 11111 } })]), slideOf(5, [shape('s5-e1', { shape: 'chevron', adj: { adj: 22222 } })])], 5)
    expect(Object.keys(c.adjust).sort()).toEqual(['1', '2'])
    const q = await openPptx(out)
    expect(gdsOf(await q.xml('ppt/slides/slide1.xml'), 'Shape 1')).toEqual([['adj', 'val 11111']])
    expect(gdsOf(await q.xml('ppt/slides/slide2.xml'), 'Shape 1')).toEqual([['adj', 'val 22222']])
  })

  it('adj は丸めてから 32 bit 整数の範囲で判定し、外は捨てて W-SHAPE', async () => {
    const values = [2147483647.4, 2147483647.5, -2147483648.4, -2147483649, 1e21]
    const { ctx: c } = await run([slideOf(1, values.map((v, i) => shape(`s1-e${i + 1}`, { shape: 'foldedCorner', adj: { adj: v } })))], 1)
    expect(c.adjust[1]).toEqual({ 'Shape 1': { adj: 2147483647 }, 'Shape 3': { adj: -2147483648 } })
    expect(c.report.warnings.filter((w) => w.code === 'W-SHAPE').map((w) => w.elementId)).toEqual(['s1-e2', 's1-e4', 's1-e5'])
  })

  it('数でない adj（文字列・真偽値・null）も捨てて、図形 1 つに W-SHAPE 1 件（範囲外と同じ文面）。残りの数は使う', async () => {
    const { ctx: c } = await run([slideOf(1, [
      shape('s1-e1', { shape: 'star5', adj: { adj: '10000', hf: 50000, vf: true, adj2: null } }),
      shape('s1-e2', { shape: 'foo', adj: { adj: 'x' } }),
    ])], 1)
    expect(c.adjust[1]).toEqual({ 'Shape 1': { hf: 50000 } })
    expect(c.report.warnings.filter((w) => w.code === 'W-SHAPE').map((w) => [w.elementId, w.message])).toEqual([
      ['s1-e1', '図形 "star5" の調整値 adj, vf, adj2 は定義に無いか、数でないか、範囲外なので捨てた'],
      ['s1-e2', '図形 "foo" は PowerPoint の図形に無いので rect にした（調整値も捨てた）'],
    ])
  })

  it('roundRect の radius の換算は定義の pin（0〜50000）に収める', async () => {
    const { ctx: c } = await run([slideOf(1, [
      shape('s1-e1', { frame: { radius: 30, inset: [0, 0, 0, 0] } }),
      shape('s1-e2', { frame: { radius: -5, inset: [0, 0, 0, 0] } }),
    ])], 1)
    expect(c.adjust[1]).toEqual({ 'Shape 1': { adj: 50000 }, 'Shape 2': { adj: 0 } })
  })

  it('線の知らない矢じりも arrow にして W-SHAPE（図形と同じ文面）。知っている値と none だけなら出さない', async () => {
    const line = (id: string, y: number, head?: string, tail?: string): LineElement => ({
      id, name: '', source: 'ppt', kind: 'line', box: { x: 10, y, w: 200, h: 0 }, boxSource: 'prop',
      from: { x: 10, y }, to: { x: 210, y }, line: { color: '#000000', width: 1, dash: 'solid', head, tail },
    })
    const { ctx: c, out } = await run([slideOf(1, [line('s1-e1', 100, 'foo', 'bar'), line('s1-e2', 200, 'stealth', 'none')])], 1)
    const ws = c.report.warnings.filter((w) => w.code === 'W-SHAPE')
    expect(ws).toHaveLength(1)
    expect(ws[0]).toMatchObject({ elementId: 's1-e1', slide: 1, message: '矢じり foo, bar は PowerPoint に無いので arrow にした' })
    const doc = await (await openPptx(out)).xml('ppt/slides/slide1.xml')
    const lines = shapesOf(doc).filter((s) => els(s, 'a', 'prstGeom')[0]?.getAttribute('prst') === 'line')
    expect(lines).toHaveLength(2)
    const ends = (s: Element) => [els(s, 'a', 'headEnd')[0]?.getAttribute('type'), els(s, 'a', 'tailEnd')[0]?.getAttribute('type')]
    expect(ends(lines[0])).toEqual(['arrow', 'arrow'])
    expect(ends(lines[1])[0]).toBe('stealth')
  })
})

function gdsOf(doc: Document, name: string): (string | null)[][] {
  return els(els(shapeNamed(doc, name)!, 'a', 'prstGeom')[0], 'a', 'gd').map((g) => [g.getAttribute('name'), g.getAttribute('fmla')])
}

describe('convert: 背景（§5.5）', () => {
  it('背景色は bg の srgbClr', () => {
    expect(els(els(slide[4], 'p', 'bg')[0], 'a', 'srgbClr')[0].getAttribute('val')).toBe('1E293B')
    expect(els(els(slide[2], 'p', 'bg')[0], 'a', 'srgbClr')[0].getAttribute('val')).toBe('FFFFFF')
  })
  it('frontmatter.background が色（#…）なら bg に入り、dim 矩形は無い', () => {
    expect(shapesOf(slide[4]).some((s) => cNvPrOf(s).getAttribute('name') === 'Background dim')).toBe(false)
    expect(ctx.shapeNames[4][0]).not.toBe('Background dim')
  })
  it('frontmatter.background が画像なら bg は blip、spTree の最初の図形が Background dim（黒、alpha 45000）', () => {
    expect(els(els(slide[1], 'p', 'bg')[0], 'a', 'blip')).toHaveLength(1)
    const first = shapesOf(slide[1])[0]
    expect(els(first, 'a', 'srgbClr')[0].getAttribute('val')).toBe('000000')
    expect(els(first, 'a', 'alpha')[0].getAttribute('val')).toBe('45000')
    expect(ctx.shapeNames[1][0]).toBe('Background dim')
    expect(ctx.report.dropped['background-crop']).toBe(1)
  })
})

describe('convert: PatchContext と Report（§4.5、§3.3、§8）', () => {
  it('shapeNames は add した順（Background dim → 要素）', () => {
    expect(ctx.shapeNames[1]).toEqual(['Background dim', 'Title', 'Body'])
    expect(ctx.shapeNames[2]).toEqual(['Title', 'Body', 'Text 3', 'Table 4', 'photo', 'arrow', 'Shape 7', 'Line 8', 'Replaced 9', 'Image 10'])
    expect(ctx.shapeNames[3]).toEqual(['Title', 'Body 2', 'Text 3'])
  })
  it('spTree の図形数は shapeNames の長さ以上で、先頭から順に対応する（§4.5 の前提。表・画像・図形が混ざる場合。§11 の 6）', () => {
    for (const no of [1, 2, 3]) {
      const shapes = shapesOf(slide[no])
      expect(shapes.length).toBeGreaterThanOrEqual(ctx.shapeNames[no].length)
      // 後処理前なので name は PptxGenJS の採番のまま。種類の並びで対応を確かめる
      const kinds = shapes.slice(0, ctx.shapeNames[no].length).map((s) => s.localName)
      const expected = ctx.shapeNames[no].map((n) => (/^(Table)/.test(n) ? 'graphicFrame' : /^(Image|Replaced|photo)/.test(n) ? 'pic' : 'sp'))
      expect(kinds).toEqual(expected)
    }
  })
  it('縮小率: contentHeight 900 > boxHeight 472 → fontScale floor(472/900×100)×1000 = 52000、lnSpcReduction 20000、W-OVERFLOW', () => {
    expect(ctx.autofit[3]['Body 2']).toEqual({ fontScale: 52000, lnSpcReduction: 20000 })
    expect(ctx.report.warnings.some((w) => w.code === 'W-OVERFLOW' && w.slide === 3)).toBe(true)
  })
  it('縮小率の下限は 25000、90000 以上なら lnSpcReduction 10000', async () => {
    const capture = structuredClone(captureJson) as unknown as Capture
    const big = capture.slides[2].elements[1] as { fit: { contentHeight: number; boxHeight: number } }
    big.fit = { contentHeight: 4720, boxHeight: 472 }
    const small = structuredClone(captureJson) as unknown as Capture
    ;(small.slides[2].elements[1] as { fit: { contentHeight: number; boxHeight: number } }).fit = { contentHeight: 500, boxHeight: 472 }
    const a = await build(capture, data(), { assets: {}, lang: 'ja-JP', layouts: dataJson.layouts })
    expect(a.ctx.autofit[3]['Body 2']).toEqual({ fontScale: 25000, lnSpcReduction: 20000 })
    const b = await build(small, data(), { assets: {}, lang: 'ja-JP', layouts: dataJson.layouts })
    expect(b.ctx.autofit[3]['Body 2']).toEqual({ fontScale: 94000, lnSpcReduction: 10000 })
  })
  it('はみ出し: x=-30 は 0 に寄せて w を縮め、W-OFFSLIDE（5 px 以上）', () => {
    const t = shapesOf(slide[3]).find((s) => textOf(s) === 'はみ出し')!
    expect(els(t, 'a', 'off')[0].getAttribute('x')).toBe('0')
    expect(Number(els(t, 'a', 'ext')[0].getAttribute('cx'))).toBe(emu(170, canvas))
    expect(ctx.report.warnings.some((w) => w.code === 'W-OFFSLIDE' && w.slide === 3)).toBe(true)
  })
  it('Report: 件数と、Capture の警告にスライド番号と名前が付く', () => {
    const all = (captureJson as unknown as Capture).slides.flatMap((s) => s.elements)
    expect(ctx.report.slides).toBe(5)
    expect(ctx.report.replaced).toBe(1)
    expect(ctx.report.native).toBe(all.length - 1)
    expect(ctx.report.replacements[0]).toMatchObject({ slide: 2, elementId: 's2-e9', name: 'Replaced 9', reason: 'math' })
    const w = ctx.report.warnings.find((x) => x.code === 'W-MATH-INLINE')!
    expect(w).toMatchObject({ slide: 2, name: 'Body' })
  })
  it('ノート: [click] を消し、行頭の - を • に', async () => {
    const notes = await p.xml('ppt/notesSlides/notesSlide1.xml')
    expect(textOf(notes as unknown as Element)).toContain('• 箇条書き')
    expect(textOf(notes as unknown as Element)).not.toContain('[click]')
  })
})

function relIdOfLayout(relsDoc: Document): string {
  return els(relsDoc, 'rel', 'Relationship').find((r) => r.getAttribute('Type')!.endsWith('/slideLayout'))!.getAttribute('Id')!
}
