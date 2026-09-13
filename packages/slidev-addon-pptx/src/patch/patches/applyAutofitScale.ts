// 後処理 4: ctx.autofit の図形名で <p:sp> を引き、<a:bodyPr> の既存の <a:normAutofit> に fontScale と
// lnSpcReduction を足す（無ければ 1 つ作る。2 つにはしない）（native-export.md §3.2、§3.3）。
// 前提にする変換: 1（name が確定している）
import type { Patch } from '../index.ts'
import { elements, children, NS } from '../zip.ts'
import { spTreeShapes } from './renameShapes.ts'

export const applyAutofitScale: Patch = {
  name: 'applyAutofitScale',
  async run(zip, ctx) {
    for (const [no, byName] of Object.entries(ctx.autofit)) {
      const path = `ppt/slides/slide${no}.xml`
      if (!zip.has(path) || Object.keys(byName).length === 0) continue
      const doc = await zip.readXml(path)
      for (const shape of spTreeShapes(doc)) {
        const name = elements(shape, 'p', 'cNvPr')[0]?.getAttribute('name') ?? ''
        const scale = byName[name]
        if (!scale) continue
        const bodyPr = elements(shape, 'a', 'bodyPr')[0]
        if (!bodyPr) continue
        for (const other of children(bodyPr).filter((c) => /^(spAutoFit|noAutofit)$/.test(c.localName))) bodyPr.removeChild(other)
        let na = children(bodyPr, 'normAutofit')[0]
        if (!na) {
          na = doc.createElementNS(NS.a, 'a:normAutofit')
          bodyPr.appendChild(na)
        }
        na.setAttribute('fontScale', String(scale.fontScale))
        na.setAttribute('lnSpcReduction', String(scale.lnSpcReduction))
      }
      zip.writeXml(path, doc)
    }
  },
}
