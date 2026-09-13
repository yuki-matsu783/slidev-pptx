// 後処理 3: dedupeParagraphProps（native-export.md §3.2、C14）
import { describe, expect, it } from 'vitest'
import { dedupeParagraphProps } from '../../packages/slidev-addon-pptx/src/patch/patches/dedupeParagraphProps'
import { openPptx, els, readFixturePptx } from '../helpers/pptx'
import { runPatches } from './helpers'

function paragraphs(doc: Document): Element[] {
  return els(doc, 'a', 'p')
}
function pPrsOf(p: Element): Element[] {
  return Array.from(p.childNodes).filter((n): n is Element => n.nodeType === 1 && (n as Element).localName === 'pPr')
}

describe('patch/dedupeParagraphProps', () => {
  it('fixture には <a:pPr> が 2 つ以上ある <a:p> がある（前提）', async () => {
    const p = await openPptx(readFixturePptx())
    const multi = paragraphs(await p.xml('ppt/slides/slide1.xml')).filter((x) => pPrsOf(x).length >= 2)
    expect(multi.length).toBeGreaterThan(0)
  })

  it('各 <a:p> の <a:pPr> が先頭の 1 つだけになる', async () => {
    const p = await openPptx(await runPatches([dedupeParagraphProps]))
    for (const path of ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/notesSlides/notesSlide1.xml']) {
      for (const para of paragraphs(await p.xml(path))) {
        const pPrs = pPrsOf(para)
        expect(pPrs.length).toBeLessThanOrEqual(1)
        if (pPrs.length === 1) {
          const first = Array.from(para.childNodes).find((n) => n.nodeType === 1) as Element
          expect(first.localName).toBe('pPr')
        }
      }
    }
  })

  it('残るのは先頭の <a:pPr>（箇条書きの marL / buChar が消えない）', async () => {
    const p = await openPptx(await runPatches([dedupeParagraphProps]))
    const doc = await p.xml('ppt/slides/slide1.xml')
    const bulleted = paragraphs(doc).filter((x) => els(x, 'a', 'buChar').length)
    expect(bulleted.length).toBeGreaterThanOrEqual(2)
    for (const para of bulleted) expect(pPrsOf(para)[0].getAttribute('marL')).toBe('342900')
  })

  it('run の数と文字は変わらない', async () => {
    const before = await openPptx(readFixturePptx())
    const after = await openPptx(await runPatches([dedupeParagraphProps]))
    const b = els(await before.xml('ppt/slides/slide1.xml'), 'a', 't').map((t) => t.textContent)
    const a = els(await after.xml('ppt/slides/slide1.xml'), 'a', 't').map((t) => t.textContent)
    expect(a).toEqual(b)
  })
})
