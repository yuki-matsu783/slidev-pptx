// フォント（native-export.md §6）
import { describe, expect, it } from 'vitest'
import { fontFor, themeFonts } from '../../packages/slidev-addon-pptx/src/build/fonts'

describe('build/fonts', () => {
  it("本文は日本語名の '游ゴシック'（PptxGenJS のテーマの Jpan 指定と同じ）", () => {
    expect(fontFor('body')).toBe('游ゴシック')
  })
  it("等幅は 'Consolas'", () => {
    expect(fontFor('code')).toBe('Consolas')
  })
  it('テーマにも同じ名前を入れる（run に fontFace の付かない箇所を拾うため）', () => {
    expect(themeFonts()).toEqual({ headFontFace: '游ゴシック', bodyFontFace: '游ゴシック' })
  })
})
