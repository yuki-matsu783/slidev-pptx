// レイアウト対応表とマスターの定義（native-export.md §5.2、§5.3）
import { describe, expect, it } from 'vitest'
import PptxGenJS from 'pptxgenjs'
import { LAYOUTS, defineMasters, masterFor } from '../../packages/slidev-addon-pptx/src/build/masters'
import { openPptx, els, shapesOf } from '../helpers/pptx'

const IN = 914400
const TOL = 0.01 * IN // 対応表は小数 2 桁なので 0.01 インチまで許す

interface Ph { name: string; type: string; x: number; y: number; w: number; h: number }
const table: Record<string, Ph[]> = {
  cover: [
    { name: 'title', type: 'title', x: 0.76, y: 2.99, w: 11.81, h: 1.09 },
    { name: 'body', type: 'body', x: 0.76, y: 4.19, w: 11.81, h: 1.5 },
  ],
  default: [
    { name: 'title', type: 'title', x: 0.76, y: 0.54, w: 11.81, h: 0.6 },
    { name: 'body', type: 'body', x: 0.76, y: 1.31, w: 11.81, h: 5.66 },
  ],
  center: [
    { name: 'title', type: 'title', x: 0.76, y: 3.21, w: 11.81, h: 0.6 },
    { name: 'body', type: 'body', x: 0.76, y: 3.97, w: 11.81, h: 1.5 },
  ],
  'two-cols': [
    { name: 'title', type: 'title', x: 0.76, y: 0.54, w: 5.9, h: 0.6 },
    { name: 'body', type: 'body', x: 0.76, y: 1.31, w: 5.9, h: 5.66 },
    { name: 'body2', type: 'body', x: 6.67, y: 0.54, w: 5.9, h: 6.42 },
  ],
}

describe('build/masters: LAYOUTS', () => {
  it('キーの順は blank, cover, default, center, two-cols（slideLayout 2〜6 の順）', () => {
    expect(Object.keys(LAYOUTS)).toEqual(['blank', 'cover', 'default', 'center', 'two-cols'])
  })
  it('title は Slidev のレイアウト名と同じ', () => {
    for (const [k, v] of Object.entries(LAYOUTS)) expect(v.title).toBe(k)
  })
  it('blank は placeholder を持たない', () => {
    expect((LAYOUTS.blank.objects ?? []).filter((o) => 'placeholder' in o)).toHaveLength(0)
  })
  it('placeholder には座標と name / type 以外を置かない（レイアウト側が呼び出し側を総取りで上書きするため）', () => {
    for (const v of Object.values(LAYOUTS)) {
      for (const o of v.objects ?? []) {
        if (!('placeholder' in o)) continue
        const keys = Object.keys(o.placeholder.options).sort()
        expect(keys).toEqual(['h', 'name', 'type', 'w', 'x', 'y'])
      }
    }
  })
  it('slideNumber は付けない（DEFAULT レイアウトに番号 placeholder が増えるため）', () => {
    for (const v of Object.values(LAYOUTS)) expect(v.slideNumber).toBeUndefined()
  })
})

describe('build/masters: defineMasters の出力', () => {
  async function written() {
    const pptx = new PptxGenJS()
    pptx.layout = 'LAYOUT_WIDE'
    defineMasters(pptx, 980)
    for (const name of Object.keys(LAYOUTS)) pptx.addSlide({ masterName: name })
    return openPptx(await pptx.write({ outputType: 'nodebuffer' }) as Buffer)
  }

  it('slideLayout の番号: 1 = DEFAULT, 2 = blank, 3 = cover, 4 = default, 5 = center, 6 = two-cols', async () => {
    const p = await written()
    const names: Record<number, string> = {}
    for (let i = 1; i <= 6; i++) {
      const doc = await p.xml(`ppt/slideLayouts/slideLayout${i}.xml`)
      names[i] = els(doc, 'p', 'cSld')[0].getAttribute('name')!
    }
    expect(names).toEqual({ 1: 'DEFAULT', 2: 'blank', 3: 'cover', 4: 'default', 5: 'center', 6: 'two-cols' })
    expect(p.list().filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f))).toHaveLength(6)
  })

  it('placeholder の座標と type が対応表どおり（0.01 インチ以内）', async () => {
    const p = await written()
    const layoutNo: Record<string, number> = { cover: 3, default: 4, center: 5, 'two-cols': 6 }
    for (const [layout, phs] of Object.entries(table)) {
      const doc = await p.xml(`ppt/slideLayouts/slideLayout${layoutNo[layout]}.xml`)
      const shapes = shapesOf(doc).filter((s) => els(s, 'p', 'ph').length)
      expect(shapes, layout).toHaveLength(phs.length)
      shapes.forEach((s, i) => {
        const ph = els(s, 'p', 'ph')[0]
        expect(ph.getAttribute('type'), `${layout}/${phs[i].name}`).toBe(phs[i].type)
        expect(ph.getAttribute('idx')).toBe(String(100 + i))
        const off = els(s, 'a', 'off')[0]
        const ext = els(s, 'a', 'ext')[0]
        expect(Math.abs(Number(off.getAttribute('x')) - phs[i].x * IN), `${layout}/${phs[i].name} x`).toBeLessThanOrEqual(TOL)
        expect(Math.abs(Number(off.getAttribute('y')) - phs[i].y * IN), `${layout}/${phs[i].name} y`).toBeLessThanOrEqual(TOL)
        expect(Math.abs(Number(ext.getAttribute('cx')) - phs[i].w * IN), `${layout}/${phs[i].name} w`).toBeLessThanOrEqual(TOL)
        expect(Math.abs(Number(ext.getAttribute('cy')) - phs[i].h * IN), `${layout}/${phs[i].name} h`).toBeLessThanOrEqual(TOL)
      })
    }
  })

  it('対応表は Slidev の px（56 px の padding など）で定義し、canvasWidth で換算する: 1960 px なら座標は半分', async () => {
    const a = new PptxGenJS(); a.layout = 'LAYOUT_WIDE'; defineMasters(a, 980); a.addSlide({ masterName: 'default' })
    const b = new PptxGenJS(); b.layout = 'LAYOUT_WIDE'; defineMasters(b, 1960); b.addSlide({ masterName: 'default' })
    const pa = await openPptx(await a.write({ outputType: 'nodebuffer' }) as Buffer)
    const pb = await openPptx(await b.write({ outputType: 'nodebuffer' }) as Buffer)
    const xa = Number(els(await pa.xml('ppt/slideLayouts/slideLayout4.xml'), 'a', 'off')[0].getAttribute('x'))
    const xb = Number(els(await pb.xml('ppt/slideLayouts/slideLayout4.xml'), 'a', 'off')[0].getAttribute('x'))
    expect(Math.abs(xb * 2 - xa)).toBeLessThanOrEqual(2)
  })

  it('defineMasters は LAYOUTS を書き換えない（PptxGenJS の createSlideMaster は渡した options を破壊するので、写しを渡す）', async () => {
    const before = JSON.stringify(LAYOUTS)
    const pptx = new PptxGenJS(); pptx.layout = 'LAYOUT_WIDE'
    defineMasters(pptx, 980)
    defineMasters(new PptxGenJS(), 980)
    expect(JSON.stringify(LAYOUTS)).toBe(before)
    for (const v of Object.values(LAYOUTS)) {
      for (const o of v.objects ?? []) {
        if ('placeholder' in o) expect(Object.keys(o.placeholder.options).sort()).toEqual(['h', 'name', 'type', 'w', 'x', 'y'])
      }
    }
  })
})

describe('build/masters: masterFor（§5.1 の 2 段目）', () => {
  it('対応表にある名前はそのまま', () => {
    for (const k of ['cover', 'default', 'center', 'two-cols']) expect(masterFor(k)).toBe(k)
  })
  it('対応表に無い名前は blank', () => {
    for (const k of ['section', 'quote', 'image-right', 'intro', 'end', 'none', 'two-cols-header']) expect(masterFor(k)).toBe('blank')
  })
})
