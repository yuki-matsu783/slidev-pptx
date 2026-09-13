// 後処理 2: dropEmptyPlaceholders（native-export.md §3.2）
import { describe, expect, it } from 'vitest'
import { dropEmptyPlaceholders } from '../../packages/slidev-addon-pptx/src/patch/patches/dropEmptyPlaceholders'
import { openPptx, shapesOf, els, textOf, readFixturePptx } from '../helpers/pptx'
import { runPatches } from './helpers'

describe('patch/dropEmptyPlaceholders', () => {
  it('文字の無い placeholder の <p:sp> だけ消える', async () => {
    const before = await openPptx(readFixturePptx())
    const s2before = shapesOf(await before.xml('ppt/slides/slide2.xml'))
    expect(s2before.filter((s) => els(s, 'p', 'ph').length).length).toBe(2)

    const after = await openPptx(await runPatches([dropEmptyPlaceholders]))
    const s2 = shapesOf(await after.xml('ppt/slides/slide2.xml'))
    expect(s2.filter((s) => els(s, 'p', 'ph').length)).toHaveLength(0)
    expect(s2).toHaveLength(2) // 自由配置テキストと線は残る
  })

  it('文字のある placeholder（title）は残る', async () => {
    const after = await openPptx(await runPatches([dropEmptyPlaceholders]))
    const s1 = shapesOf(await after.xml('ppt/slides/slide1.xml'))
    const ph = s1.filter((s) => els(s, 'p', 'ph').length)
    expect(ph).toHaveLength(1)
    expect(textOf(ph[0])).toBe('見出し')
  })

  it('レイアウトの placeholder は触らない', async () => {
    const before = await openPptx(readFixturePptx())
    const after = await openPptx(await runPatches([dropEmptyPlaceholders]))
    expect(await after.text('ppt/slideLayouts/slideLayout2.xml')).toBe(await before.text('ppt/slideLayouts/slideLayout2.xml'))
  })
})
