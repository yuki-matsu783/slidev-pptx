// 後処理 4 の直後: ctx.adjust の図形名で <p:sp> を引き、<a:prstGeom> の <a:avLst> の中身を <a:gd name fmla="val N"/> で
// 置き換える（PptxGenJS は avLst を rectRadius と angleRange でしか書けない）。並びは定義の avLst の順で、
// 定義に無い名前は書かない（変換で捨てて W-SHAPE を出してある）。
// 前提にする変換: 1（name が確定している）
import type { Patch } from '../index.ts'
import { elements, children, NS } from '../zip.ts'
import { spTreeShapes } from './renameShapes.ts'
import { PRESETS } from '../../shapes/presets.ts'

export const applyShapeAdjust: Patch = {
  name: 'applyShapeAdjust',
  async run(zip, ctx) {
    for (const [no, byName] of Object.entries(ctx.adjust)) {
      const path = `ppt/slides/slide${no}.xml`
      if (!zip.has(path) || Object.keys(byName).length === 0) continue
      const doc = await zip.readXml(path)
      for (const shape of spTreeShapes(doc)) {
        if (shape.localName !== 'sp') continue
        const name = elements(shape, 'p', 'cNvPr')[0]?.getAttribute('name') ?? ''
        const values = byName[name]
        if (!values) continue
        const geom = elements(shape, 'a', 'prstGeom')[0]
        const prst = geom?.getAttribute('prst') ?? ''
        if (!geom || !Object.hasOwn(PRESETS, prst)) continue
        const def = PRESETS[prst]
        let avLst = children(geom, 'avLst')[0]
        if (!avLst) {
          avLst = doc.createElementNS(NS.a, 'a:avLst')
          geom.insertBefore(avLst, geom.firstChild)
        }
        for (const old of children(avLst)) avLst.removeChild(old)
        for (const [key] of def.av) {
          if (!Object.hasOwn(values, key)) continue
          const gd = doc.createElementNS(NS.a, 'a:gd')
          gd.setAttribute('name', key)
          gd.setAttribute('fmla', `val ${Math.round(values[key])}`)
          avLst.appendChild(gd)
        }
      }
      zip.writeXml(path, doc)
    }
  },
}
