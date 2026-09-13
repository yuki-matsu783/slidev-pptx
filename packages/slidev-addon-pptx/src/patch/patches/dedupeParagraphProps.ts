// 後処理 3: 各 <a:p> の <a:pPr> を先頭の 1 つだけ残す（native-export.md §3.2、C14）。
// PptxGenJS は run ごとに <a:pPr> を出し、bullet の無い run にも <a:buNone/> 付きの <a:pPr> を書く。
// 前提にする変換: なし
import { xmlPatch } from '../index.ts'
import { elements, children } from '../zip.ts'

export const dedupeParagraphProps = xmlPatch('dedupeParagraphProps', /^ppt\/(slides\/slide\d+|notesSlides\/notesSlide\d+)\.xml$/, (doc) => {
  for (const p of elements(doc, 'a', 'p')) {
    const pPrs = children(p, 'pPr')
    if (pPrs.length === 0) continue
    const [first, ...rest] = pPrs
    for (const extra of rest) p.removeChild(extra)
    if (p.firstChild !== first) {
      p.removeChild(first)
      p.insertBefore(first, p.firstChild)
    }
  }
})
