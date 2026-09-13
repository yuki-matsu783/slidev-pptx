// 後処理 5: splitNotesParagraphs（native-export.md §3.2、§4.4）
import { describe, expect, it } from 'vitest'
import { splitNotesParagraphs } from '../../packages/slidev-addon-pptx/src/patch/patches/splitNotesParagraphs'
import { openPptx, els, readFixturePptx } from '../helpers/pptx'
import { runPatches } from './helpers'

function bodyParagraphs(doc: Document): Element[] {
  const body = els(doc, 'p', 'sp').find((sp) => els(sp, 'p', 'ph')[0]?.getAttribute('type') === 'body')!
  return els(body, 'a', 'p')
}

describe('patch/splitNotesParagraphs', () => {
  it('fixture のノートは 1 つの <a:t> に CRLF で詰まっている（前提）', async () => {
    const p = await openPptx(readFixturePptx())
    const doc = await p.xml('ppt/notesSlides/notesSlide1.xml')
    const ts = els(doc, 'a', 't').map((t) => t.textContent)
    expect(ts.some((t) => /\r\n/.test(t ?? ''))).toBe(true)
  })

  it('行ごとに <a:p> に割れる', async () => {
    const p = await openPptx(await runPatches([splitNotesParagraphs]))
    const paras = bodyParagraphs(await p.xml('ppt/notesSlides/notesSlide1.xml'))
    const texts = paras.map((para) => els(para, 'a', 't').map((t) => t.textContent).join(''))
    expect(texts).toEqual(['1 行目', '2 行目', '3 行目'])
    expect(texts.join('')).not.toMatch(/\r|\n/)
  })

  it('割った <a:p> の <a:pPr> は先頭に高々 1 つ（C14 を保つ）', async () => {
    const p = await openPptx(await runPatches([splitNotesParagraphs]))
    for (const para of bodyParagraphs(await p.xml('ppt/notesSlides/notesSlide1.xml'))) {
      const kids = Array.from(para.childNodes).filter((n) => n.nodeType === 1) as Element[]
      expect(kids.filter((k) => k.localName === 'pPr').length).toBeLessThanOrEqual(1)
      if (kids[0]?.localName !== 'pPr') expect(kids.some((k) => k.localName === 'pPr')).toBe(false)
    }
  })

  it('ノートの無いスライド（notesSlide2）は壊れない', async () => {
    const p = await openPptx(await runPatches([splitNotesParagraphs]))
    const doc = await p.xml('ppt/notesSlides/notesSlide2.xml')
    expect(els(doc, 'p', 'notes')).toHaveLength(1)
  })
})
