// フォント（native-export.md §6）。名前の付与はここに集める。案 b（テーマ任せ）に切り替えるときは fontFor が undefined を返す。

export const BODY_FONT = '游ゴシック'
export const CODE_FONT = 'Consolas'

export function fontFor(kind: 'body' | 'code'): string {
  return kind === 'code' ? CODE_FONT : BODY_FONT
}

export function themeFonts(): { headFontFace: string; bodyFontFace: string } {
  return { headFontFace: BODY_FONT, bodyFontFace: BODY_FONT }
}
