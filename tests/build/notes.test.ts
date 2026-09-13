// 発表者ノートの前処理（native-export.md §4.4）
import { describe, expect, it } from 'vitest'
import { notesText } from '../../packages/slidev-addon-pptx/src/build/notes'

describe('build/notes: notesText', () => {
  it('[click] と [click:N] を消す（前後の空白も詰める）', () => {
    expect(notesText('1 行目\n[click] 2 行目\n[click:3] 3 行目')).toBe('1 行目\n2 行目\n3 行目')
  })
  it('行頭の - と * を • に', () => {
    expect(notesText('- a\n* b\n  - c')).toBe('• a\n• b\n  • c')
  })
  it('** _ ` の記号はそのまま残す（Markdown の変換器を持ち込まない）', () => {
    expect(notesText('**太字** と _斜体_ と `code`')).toBe('**太字** と _斜体_ と `code`')
  })
  it('空のノートは空文字', () => {
    expect(notesText('')).toBe('')
    expect(notesText(undefined)).toBe('')
  })
  it('改行は \\n のまま（CRLF への置き換えと段落分けは PptxGenJS と後処理 5 がやる）', () => {
    expect(notesText('a\r\nb')).toBe('a\nb')
  })
})
