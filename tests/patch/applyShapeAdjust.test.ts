// 後処理: applyShapeAdjust（図形の調整値。<a:avLst> を定義の avLst の順で書き直す）
import { describe, expect, it } from 'vitest'
import PptxGenJS from 'pptxgenjs'
import { renameShapes } from '../../packages/slidev-addon-pptx/src/patch/patches/renameShapes'
import { applyShapeAdjust } from '../../packages/slidev-addon-pptx/src/patch/patches/applyShapeAdjust'
import type { PatchContext } from '../../packages/slidev-addon-pptx/src/patch/index'
import { openPptx, shapeNamed, els } from '../helpers/pptx'
import { contextFor, runPatches } from './helpers'

// sample.pptx には調整値のある図形が無いので、PptxGenJS で小さな入力を作る
async function input(): Promise<Buffer> {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'
  const s = pptx.addSlide()
  s.addShape('roundRect' as PptxGenJS.SHAPE_NAME, { x: 1, y: 1, w: 2, h: 1, rectRadius: 0.2, objectName: 'Card' })
  s.addShape('bentConnector3' as PptxGenJS.SHAPE_NAME, { x: 1, y: 3, w: 2, h: 1, objectName: 'Conn' })
  s.addShape('roundRect' as PptxGenJS.SHAPE_NAME, { x: 4, y: 1, w: 2, h: 1, rectRadius: 0.2, objectName: 'Keep' })
  s.addText('文字', { shape: 'wedgeRectCallout' as PptxGenJS.SHAPE_NAME, x: 4, y: 3, w: 2, h: 1, objectName: 'Callout' })
  return (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
}

const ctxWith = (adjust: PatchContext['adjust']) => contextFor({ shapeNames: { 1: ['Card', 'Conn', 'Keep', 'Callout'] }, adjust })
const gds = (slide: Document, name: string) =>
  els(els(shapeNamed(slide, name)!, 'a', 'prstGeom')[0], 'a', 'gd').map((g) => [g.getAttribute('name'), g.getAttribute('fmla')])

describe('patch/applyShapeAdjust', () => {
  it('既存の avLst の gd（rectRadius の adj）を調整値で置き換える', async () => {
    const buf = await input()
    const before = await openPptx(await runPatches([renameShapes], ctxWith({}), buf))
    expect(gds(await before.xml('ppt/slides/slide1.xml'), 'Card')).toHaveLength(1)
    const p = await openPptx(await runPatches([renameShapes, applyShapeAdjust], ctxWith({ 1: { Card: { adj: 30000 } } }), buf))
    expect(gds(await p.xml('ppt/slides/slide1.xml'), 'Card')).toEqual([['adj', 'val 30000']])
  })

  it('空の avLst に gd を入れ、並びは定義の avLst の順、値は整数に丸める', async () => {
    const ctx = ctxWith({ 1: { Conn: { adj1: -20000 }, Callout: { adj2: 12345.6, adj1: 100 } } })
    const slide = await (await openPptx(await runPatches([renameShapes, applyShapeAdjust], ctx, await input()))).xml('ppt/slides/slide1.xml')
    expect(gds(slide, 'Conn')).toEqual([['adj1', 'val -20000']])
    expect(gds(slide, 'Callout')).toEqual([['adj1', 'val 100'], ['adj2', 'val 12346']])
  })

  it('定義に無い名前は書かない', async () => {
    const ctx = ctxWith({ 1: { Conn: { adj1: 1, adj9: 2 } } })
    const slide = await (await openPptx(await runPatches([renameShapes, applyShapeAdjust], ctx, await input()))).xml('ppt/slides/slide1.xml')
    expect(gds(slide, 'Conn')).toEqual([['adj1', 'val 1']])
  })

  it('名前が合わない図形は触らない', async () => {
    const buf = await input()
    const before = await openPptx(await runPatches([renameShapes], ctxWith({}), buf))
    const p = await openPptx(await runPatches([renameShapes, applyShapeAdjust], ctxWith({ 1: { Card: { adj: 30000 }, Nothing: { adj: 1 } } }), buf))
    const slide = await p.xml('ppt/slides/slide1.xml')
    expect(gds(slide, 'Keep')).toEqual(gds(await before.xml('ppt/slides/slide1.xml'), 'Keep'))
    expect(gds(slide, 'Conn')).toEqual([])
  })

  it('adjust が空なら何もしない（スライドの XML が 1 変換目だけのときと同じ）', async () => {
    const buf = await input()
    const once = await openPptx(await runPatches([renameShapes], ctxWith({}), buf))
    const p = await openPptx(await runPatches([renameShapes, applyShapeAdjust], ctxWith({ 1: {} }), buf))
    expect(await p.text('ppt/slides/slide1.xml')).toBe(await once.text('ppt/slides/slide1.xml'))
  })

  it('往復で冪等（2 回通しても 1 回と同じ）', async () => {
    const ctx = ctxWith({ 1: { Card: { adj: 30000 }, Callout: { adj1: 100, adj2: 200 } } })
    const once = await runPatches([renameShapes, applyShapeAdjust], ctx, await input())
    const twice = await runPatches([renameShapes, applyShapeAdjust], ctx, once)
    expect(await (await openPptx(twice)).text('ppt/slides/slide1.xml')).toBe(await (await openPptx(once)).text('ppt/slides/slide1.xml'))
  })
})
