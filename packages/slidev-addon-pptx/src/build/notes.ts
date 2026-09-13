// 発表者ノートの前処理（native-export.md §4.4）。Markdown の変換器は持ち込まない。

export function notesText(note: string | undefined): string {
  return (note ?? '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.replace(/\[click(?::\d+)?\]\s*/g, '').replace(/^(\s*)[-*]\s+/, '$1• '))
    .join('\n')
}
