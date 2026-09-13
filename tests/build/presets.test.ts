// 図形 187 種すべてを build → 後処理 → OPC 検査に通す。Capture は合成（ブラウザ無し）。
// 1 枚目は文字の無い図形（addShape）、2 枚目は文字のある図形（addText）。どちらも調整値・反転・矢じり付き
import { beforeAll, describe, expect, it } from 'vitest'
import { build } from '../../packages/slidev-addon-pptx/src/build/convert'
import type { PatchContext } from '../../packages/slidev-addon-pptx/src/patch/index'
import { PATCHES, postProcess } from '../../packages/slidev-addon-pptx/src/patch/index'
import { check } from '../../packages/slidev-addon-pptx/src/opc/check'
import { PRESET_NAMES } from '../../packages/slidev-addon-pptx/src/shapes/geometry'
import { PRESETS } from '../../packages/slidev-addon-pptx/src/shapes/presets'
import type { Capture, DeckData, ShapeElement } from '../../packages/slidev-addon-pptx/src/types'
import { openPptx, els, shapesOf, cNvPrOf } from '../helpers/pptx'

const COLS = 14
const CW = 70
const CH = 39

/** 定義の既定値（avLst の `val N`）。調整値の往復を見るため、既定と違う値にずらして渡す */
const shifted = (shape: string) => Object.fromEntries(PRESETS[shape].av.map(([n, f]) => [n, Number(f.split(' ')[1]) + 1]))

function shapes(no: number, withText: boolean): ShapeElement[] {
  return PRESET_NAMES.map((shape, i) => ({
    id: `s${no}-e${i + 1}`,
    name: '',
    source: 'ppt',
    kind: 'shape',
    shape,
    box: { x: (i % COLS) * CW + 2, y: Math.floor(i / COLS) * CH + 2, w: CW - 4, h: CH - 4 },
    boxSource: 'prop',
    frame: { fill: { color: '#3b82f6' }, line: { color: '#1d4ed8', width: 1, dash: 'solid' }, inset: [0, 0, 0, 0] },
    adj: shifted(shape),
    flipH: i % 2 === 0,
    flipV: i % 3 === 0,
    arrow: { head: 'triangle', tail: 'arrow' },
    paragraphs: withText
      ? [{ kind: 'plain', level: 0, align: 'center', lineHeight: 12, spaceBefore: 0, spaceAfter: 0, runs: [{ text: shape, size: 8, color: '#ffffff', bold: false, italic: false, underline: false, strike: false, code: false }] }]
      : undefined,
    valign: 'middle',
  }))
}

const capture: Capture = {
  canvas: { width: 980, height: 552 },
  slides: [
    { no: 1, lang: 'ja-JP', zoom: 1, backgroundColor: '#ffffff', elements: shapes(1, false), warnings: [] },
    { no: 2, lang: 'ja-JP', zoom: 1, backgroundColor: '#ffffff', elements: shapes(2, true), warnings: [] },
  ],
}
const data: DeckData = { slides: [{ index: 0, frontmatter: { layout: 'default' } }, { index: 1, frontmatter: {} }], layouts: ['default'] }

let ctx: PatchContext
let raw: Buffer
let out: Buffer

beforeAll(async () => {
  const r = await build(capture, data, { assets: {}, lang: 'ja-JP', layouts: ['default'] })
  ctx = r.ctx
  raw = (await r.pptx.write({ outputType: 'nodebuffer' })) as Buffer
  out = await postProcess(raw, PATCHES, ctx)
})

const shapeSps = (slide: Document) => shapesOf(slide).filter((s) => /^Shape \d+$/.test(cNvPrOf(s).getAttribute('name') ?? ''))
const prst = (sp: Element) => els(sp, 'a', 'prstGeom')[0]?.getAttribute('prst')

describe('build: 図形 187 種', () => {
  it('W-SHAPE が出ない（名前も調整値も全部定義にある）', () => {
    expect(ctx.report.warnings.filter((w) => w.code === 'W-SHAPE')).toEqual([])
  })

  it('PptxGenJS の出力で、各スライドの prst が 187 種そろう（ShapeType に無い名前もそのまま出る）', async () => {
    const p = await openPptx(raw)
    for (const no of [1, 2]) {
      const names = shapeSps(await p.xml(`ppt/slides/slide${no}.xml`)).map(prst)
      expect(names, `slide${no}`).toEqual([...PRESET_NAMES])
    }
    expect(PRESET_NAMES).toHaveLength(187)
  })

  it('後処理のあと、avLst は定義の名前の順に調整値の gd を持つ', async () => {
    const p = await openPptx(out)
    for (const no of [1, 2]) {
      const sps = shapeSps(await p.xml(`ppt/slides/slide${no}.xml`))
      expect(sps).toHaveLength(187)
      for (const sp of sps) {
        const name = prst(sp)!
        const gd = els(els(sp, 'a', 'prstGeom')[0], 'a', 'gd').map((g) => [g.getAttribute('name'), g.getAttribute('fmla')])
        expect(gd, `slide${no} ${name}`).toEqual(Object.entries(shifted(name)).map(([k, v]) => [k, `val ${v}`]))
      }
    }
  })

  it('後処理のあとの OPC 検査は error 0', async () => {
    const results = await check(out)
    expect(results.filter((r) => r.level === 'error')).toEqual([])
  })
})
