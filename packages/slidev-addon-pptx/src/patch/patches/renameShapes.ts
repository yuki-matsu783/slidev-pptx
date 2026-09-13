// 後処理 1: スライド・レイアウト・マスターの <p:cNvPr id> を文書順に 2 から振り直し、スライドでは name を
// ctx.shapeNames の列で付け直す（native-export.md §3.2、§4.5）。ノート（notesSlides）は対象外。
// 前提にする変換: なし（最初に走る）
import type { Patch } from '../index.ts'
import { elements, children } from '../zip.ts'

const SHAPE_TAGS = new Set(['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'])

export function spTreeShapes(doc: Document): Element[] {
  const tree = elements(doc, 'p', 'spTree')[0]
  if (!tree) return []
  return children(tree).filter((n) => SHAPE_TAGS.has(n.localName))
}

function cNvPrOf(shape: Element): Element | undefined {
  // 直下の nv*Pr の中の cNvPr（グループの中の図形の cNvPr を拾わない）
  for (const nv of children(shape)) {
    if (/^nv.*Pr$/.test(nv.localName)) {
      const c = children(nv, 'cNvPr')[0]
      if (c) return c
    }
  }
  return undefined
}

function renumber(doc: Document, names?: string[]): void {
  const tree = elements(doc, 'p', 'spTree')[0]
  if (!tree) return
  const grp = children(tree, 'nvGrpSpPr')[0]
  const grpC = grp && children(grp, 'cNvPr')[0]
  if (grpC) grpC.setAttribute('id', '1')
  const shapes = spTreeShapes(doc)
  shapes.forEach((shape, i) => {
    const c = cNvPrOf(shape)
    if (!c) return
    c.setAttribute('id', String(i + 2))
    if (names) c.setAttribute('name', names[i] ?? `Placeholder ${i + 1}`)
  })
  // グループの中の図形も、スライド内で一意になるよう続きの番号を振る
  let next = shapes.length + 2
  for (const inner of elements(doc, 'p', 'cNvPr')) {
    if (inner === grpC) continue
    if (shapes.some((s) => cNvPrOf(s) === inner)) continue
    inner.setAttribute('id', String(next++))
  }
}

export const renameShapes: Patch = {
  name: 'renameShapes',
  async run(zip, ctx) {
    for (const path of zip.list()) {
      const slide = /^ppt\/slides\/slide(\d+)\.xml$/.exec(path)
      if (slide) {
        const doc = await zip.readXml(path)
        renumber(doc, ctx.shapeNames[Number(slide[1])])
        zip.writeXml(path, doc)
      } else if (/^ppt\/(slideLayouts\/slideLayout\d+|slideMasters\/slideMaster\d+)\.xml$/.test(path)) {
        const doc = await zip.readXml(path)
        renumber(doc)
        zip.writeXml(path, doc)
      }
    }
  },
}
