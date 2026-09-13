// レイアウト名の解決 1 段目（native-export.md §5.1）
import { describe, expect, it } from 'vitest'
import { resolveLayout } from '../../packages/slidev-addon-pptx/src/build/layout'
import data from '../fixtures/capture/basic.data.json'

const layouts = data.layouts

describe('build/layout: resolveLayout（Slidev と同じ順で解く）', () => {
  it('frontmatter.layout があればそれ', () => {
    expect(resolveLayout(0, data, layouts)).toBe('cover')
    expect(resolveLayout(2, data, layouts)).toBe('two-cols')
    expect(resolveLayout(3, data, layouts)).toBe('section')
  })
  it('無ければ 1 枚目は cover、それ以外は default', () => {
    const d = { ...data, slides: [{ index: 0, frontmatter: {} }, { index: 1, frontmatter: {} }] }
    expect(resolveLayout(0, d, layouts)).toBe('cover')
    expect(resolveLayout(1, d, layouts)).toBe('default')
  })
  it('slides[0].frontmatter.defaults.layout は 1 枚目を含む全スライドの既定になる（Slidev の serve-*.mjs L683: (layout ?? defaults.layout) || (index===0 ? cover : default)）', () => {
    const d = { ...data, slides: [{ index: 0, frontmatter: { defaults: { layout: 'center' } } }, { index: 1, frontmatter: {} }] }
    expect(resolveLayout(1, d, layouts)).toBe('center')
    expect(resolveLayout(0, d, layouts)).toBe('center')
  })
  it('getLayouts() に無い名前は default に落とす（Slidev が default で描くため）', () => {
    expect(resolveLayout(4, data, layouts)).toBe('default')
  })
  it('getLayouts() に無い既定（defaults.layout の綴り違い）も default', () => {
    const d = { ...data, slides: [{ index: 0, frontmatter: { defaults: { layout: 'nope' } } }, { index: 1, frontmatter: {} }] }
    expect(resolveLayout(1, d, layouts)).toBe('default')
  })
})
