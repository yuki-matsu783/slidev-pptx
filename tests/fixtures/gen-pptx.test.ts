// fixture の PPTX が、後処理の検査が前提にする 4 条件を満たし続けていることを固定する。
// 設計: native-export.md §3.2、§11 の 6・7（PptxGenJS の版を上げたときの確認）
import { describe, expect, it } from 'vitest'
import { assertFixtureConditions, generate } from './gen-pptx.mjs'
import { openPptx, readFixturePptx, shapesOf, cNvPrOf, els } from '../helpers/pptx'

describe('fixtures/gen-pptx', () => {
  it('生成した PPTX が 4 条件（id 重複 / 空 placeholder / 実在しない Override / pPr 重複）を満たす', async () => {
    const buf = await generate()
    expect(await assertFixtureConditions(buf)).toEqual([])
  })

  it('コミット済みの sample.pptx も同じ 4 条件を満たす（生成し直し忘れの検知）', async () => {
    expect(await assertFixtureConditions(readFixturePptx())).toEqual([])
  })

  it('spTree の順は add した順で、自動追加の placeholder は末尾（§4.5 / §11 の 6）', async () => {
    const p = await openPptx(readFixturePptx())
    const s1 = await p.xml('ppt/slides/slide1.xml')
    const names = shapesOf(s1).map((s) => cNvPrOf(s).getAttribute('name'))
    // add した順: title(placeholder) → 自由配置テキスト → 表 → 画像。その後に自動追加された body
    expect(names.slice(0, 4)).toEqual(['Text 0', 'free-text', 'Table 0', 'Image 0'])
    const last = shapesOf(s1).at(-1)!
    expect(els(last, 'p', 'ph')).toHaveLength(1)
    expect(els(last, 'a', 't').map((t) => t.textContent).join('')).toBe('')
  })

  it('テキスト枠の margin は [l, r, b, t] の順（§4.2、§11 の 7）', async () => {
    // generate() と同じ API で margin だけ変えた最小の deck を作って確かめる
    const { createRequire } = await import('node:module')
    const require = createRequire(import.meta.url)
    const PptxGenJS = require('pptxgenjs')
    const pptx = new PptxGenJS()
    pptx.addSlide().addText('m', { x: 1, y: 1, w: 2, h: 1, margin: [10, 20, 30, 40] })
    const p = await openPptx(await pptx.write({ outputType: 'nodebuffer' }))
    const bodyPr = els(await p.xml('ppt/slides/slide1.xml'), 'a', 'bodyPr')[0]
    expect(bodyPr.getAttribute('lIns')).toBe(String(10 * 12700))
    expect(bodyPr.getAttribute('rIns')).toBe(String(20 * 12700))
    expect(bodyPr.getAttribute('bIns')).toBe(String(30 * 12700))
    expect(bodyPr.getAttribute('tIns')).toBe(String(40 * 12700))
  })
})
