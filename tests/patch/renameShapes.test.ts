// 後処理 1: renameShapes（native-export.md §3.2、§4.5）
import { describe, expect, it } from 'vitest'
import { renameShapes } from '../../packages/slidev-addon-pptx/src/patch/patches/renameShapes'
import { openPptx, shapesOf, cNvPrOf, els } from '../helpers/pptx'
import { contextFor, runPatches } from './helpers'

describe('patch/renameShapes', () => {
  it('スライドの cNvPr id を文書順に 2 から振り直す（1 は nvGrpSpPr 用に空ける）', async () => {
    const p = await openPptx(await runPatches([renameShapes]))
    for (const path of ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml']) {
      const doc = await p.xml(path)
      const ids = shapesOf(doc).map((s) => Number(cNvPrOf(s).getAttribute('id')))
      expect(ids).toEqual(ids.map((_, i) => i + 2))
      const grp = els(els(doc, 'p', 'spTree')[0], 'p', 'nvGrpSpPr')[0]
      expect(els(grp, 'p', 'cNvPr')[0].getAttribute('id')).toBe('1')
    }
  })

  it('name は ctx.shapeNames の列で付け直し、列より多い図形は Placeholder <位置>', async () => {
    const p = await openPptx(await runPatches([renameShapes]))
    const s1 = shapesOf(await p.xml('ppt/slides/slide1.xml')).map((s) => cNvPrOf(s).getAttribute('name'))
    // add 順 4 つ + 自動追加された空 body
    expect(s1).toEqual(['Title', 'Text 2', 'Table 3', 'Image 4', 'Placeholder 5'])
    const s2 = shapesOf(await p.xml('ppt/slides/slide2.xml')).map((s) => cNvPrOf(s).getAttribute('name'))
    expect(s2.slice(0, 2)).toEqual(['Text 1', 'Line 2'])
    expect(s2.slice(2).every((n) => /^Placeholder \d+$/.test(n!))).toBe(true)
  })

  it('スライド内で name が一意になる', async () => {
    const p = await openPptx(await runPatches([renameShapes]))
    const names = shapesOf(await p.xml('ppt/slides/slide1.xml')).map((s) => cNvPrOf(s).getAttribute('name'))
    expect(new Set(names).size).toBe(names.length)
  })

  it('レイアウトとマスターの id も一意にする（name は触らない）', async () => {
    const p = await openPptx(await runPatches([renameShapes]))
    for (const path of ['ppt/slideLayouts/slideLayout2.xml', 'ppt/slideMasters/slideMaster1.xml']) {
      const ids = shapesOf(await p.xml(path)).map((s) => cNvPrOf(s).getAttribute('id'))
      expect(new Set(ids).size).toBe(ids.length)
      expect(ids).not.toContain('0')
      expect(ids).not.toContain('1')
    }
  })

  it('ノート（notesSlides）は対象外で、元の id のまま', async () => {
    const before = await openPptx((await import('../helpers/pptx')).readFixturePptx())
    const after = await openPptx(await runPatches([renameShapes]))
    expect(await after.text('ppt/notesSlides/notesSlide1.xml')).toBe(await before.text('ppt/notesSlides/notesSlide1.xml'))
  })

  it('shapeNames に無いスライドは名前を触らず id だけ振り直す', async () => {
    const p = await openPptx(await runPatches([renameShapes], contextFor({ shapeNames: {} })))
    const names = shapesOf(await p.xml('ppt/slides/slide1.xml')).map((s) => cNvPrOf(s).getAttribute('name'))
    expect(names[1]).toBe('free-text')
  })
})
