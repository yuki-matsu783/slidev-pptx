// 後処理 1: renameShapes（native-export.md §3.2、§4.5）
import { describe, expect, it } from 'vitest'
import { renameShapes } from '../../packages/slidev-addon-pptx/src/patch/patches/renameShapes'
import { openPptx, shapesOf, cNvPrOf, els, readFixturePptx } from '../helpers/pptx'
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

  it('レイアウトの id も一意にする（fixture は元から一意なので、重複させた入力を作って通す）。name は触らない', async () => {
    const { default: JSZip } = await import('jszip')
    const zip = await JSZip.loadAsync(readFixturePptx())
    const path = 'ppt/slideLayouts/slideLayout2.xml'
    const xml = await zip.file(path)!.async('string')
    expect(xml).toMatch(/<p:cNvPr id="3"/)
    zip.file(path, xml.replace('<p:cNvPr id="3"', '<p:cNvPr id="2"')) // title と body を同じ id に
    const input = await zip.generateAsync({ type: 'nodebuffer' })
    const beforeNames = shapesOf(await (await openPptx(input)).xml(path)).map((s) => cNvPrOf(s).getAttribute('name'))

    const p = await openPptx(await runPatches([renameShapes], contextFor(), input))
    const shapes = shapesOf(await p.xml(path))
    const ids = shapes.map((s) => Number(cNvPrOf(s).getAttribute('id')))
    expect(ids).toEqual(ids.map((_, i) => i + 2))
    expect(shapes.map((s) => cNvPrOf(s).getAttribute('name'))).toEqual(beforeNames)
  })

  it('マスターの id も一意にする（fixture のマスターは図形 0 個なので、図形を 2 つ差し込んだ入力を作って通す）', async () => {
    const { default: JSZip } = await import('jszip')
    const zip = await JSZip.loadAsync(readFixturePptx())
    const path = 'ppt/slideMasters/slideMaster1.xml'
    const xml = await zip.file(path)!.async('string')
    const sp = (id: number, name: string) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="ja-JP"/><a:t>m</a:t></a:r></a:p></p:txBody></p:sp>`
    zip.file(path, xml.replace('</p:spTree>', `${sp(7, 'Logo')}${sp(7, 'Footer')}</p:spTree>`))
    const input = await zip.generateAsync({ type: 'nodebuffer' })

    const p = await openPptx(await runPatches([renameShapes], contextFor(), input))
    const shapes = shapesOf(await p.xml(path))
    expect(shapes.map((s) => Number(cNvPrOf(s).getAttribute('id')))).toEqual([2, 3])
    expect(shapes.map((s) => cNvPrOf(s).getAttribute('name'))).toEqual(['Logo', 'Footer'])
  })

  it('ノート（notesSlides）は対象外で、元の id のまま', async () => {
    const before = await openPptx(readFixturePptx())
    const after = await openPptx(await runPatches([renameShapes]))
    expect(await after.text('ppt/notesSlides/notesSlide1.xml')).toBe(await before.text('ppt/notesSlides/notesSlide1.xml'))
  })

  it('shapeNames に無いスライドは名前を触らず id だけ振り直す', async () => {
    const p = await openPptx(await runPatches([renameShapes], contextFor({ shapeNames: {} })))
    const names = shapesOf(await p.xml('ppt/slides/slide1.xml')).map((s) => cNvPrOf(s).getAttribute('name'))
    expect(names[1]).toBe('free-text')
  })
})
