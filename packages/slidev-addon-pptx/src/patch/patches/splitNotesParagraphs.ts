// 後処理 5: ノートの <a:t> の改行（CRLF / CR / LF。DOM では行末正規化で LF）を <a:p> の区切りに割る
// （native-export.md §3.2、§4.4）。割って作る <a:p> には元の <a:pPr> を先頭に 1 つだけ写す（C14 を保つ）。
// 前提にする変換: 3（<a:pPr> は先頭に高々 1 つ）
import { xmlPatch } from '../index.ts'
import { elements, children } from '../zip.ts'

export const splitNotesParagraphs = xmlPatch('splitNotesParagraphs', /^ppt\/notesSlides\/notesSlide\d+\.xml$/, (doc) => {
  for (const p of Array.from(elements(doc, 'a', 'p'))) {
    const ts = elements(p, 'a', 't')
    if (ts.length !== 1) continue
    const text = ts[0].textContent ?? ''
    const lines = text.split(/\r\n|\r|\n/)
    if (lines.length < 2) continue
    const parent = p.parentNode!
    const pPr = children(p, 'pPr')[0]
    const r = ts[0].parentNode as Element // <a:r>
    const endParaRPr = children(p, 'endParaRPr')[0]
    const next = p.nextSibling
    parent.removeChild(p)
    for (const line of lines) {
      const np = doc.createElementNS(p.namespaceURI, p.nodeName)
      if (pPr) np.appendChild(pPr.cloneNode(true))
      const nr = r.cloneNode(true) as Element
      const nt = elements(nr, 'a', 't')[0]
      while (nt.firstChild) nt.removeChild(nt.firstChild)
      nt.appendChild(doc.createTextNode(line))
      np.appendChild(nr)
      if (endParaRPr) np.appendChild(endParaRPr.cloneNode(true))
      parent.insertBefore(np, next)
    }
  }
})
