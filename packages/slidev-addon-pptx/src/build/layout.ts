// レイアウト名の解決 1 段目（native-export.md §5.1）。Slidev と同じ順で解く:
// (frontmatter.layout ?? slides[0].frontmatter.defaults.layout) || (index === 0 ? 'cover' : 'default')
// → getLayouts() に無ければ 'default'（Slidev が default で描くため）。
// 出典: @slidev/cli dist/serve-*.mjs L683-687
import type { DeckData } from '../types.ts'

export function resolveLayout(slideIndex: number, data: DeckData, layouts: string[]): string {
  const own = data.slides[slideIndex]?.frontmatter?.layout
  const defaults = data.slides[0]?.frontmatter?.defaults?.layout
  let name = (own ?? defaults) || (slideIndex === 0 ? 'cover' : 'default')
  if (!layouts.includes(name)) name = 'default'
  return name
}
