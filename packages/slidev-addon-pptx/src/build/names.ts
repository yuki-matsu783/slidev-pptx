// 図形の名前（native-export.md §4.5）。Node 側が add する順に確定し、後処理 1 が <p:cNvPr name> に付け直す。
import type { Element, RoleHint } from '../types.ts'
import { sanitizeXmlText } from './sanitize.ts'

export const BACKGROUND_DIM = 'Background dim'

const FIXED: Record<RoleHint, string> = { title: 'Title', body: 'Body', body2: 'Body 2' }

function kindName(e: Element): string {
  switch (e.kind) {
    case 'text':
      return 'Text'
    case 'table':
      return 'Table'
    case 'shape':
      return 'Shape'
    case 'line':
      return 'Line'
    case 'image':
      return e.source === 'replaced' ? 'Replaced' : 'Image'
  }
}

function elementNumber(id: string): string {
  const m = /-e(\d+)$/.exec(id)
  return m ? m[1] : id
}

/**
 * @param elements  add する順の要素
 * @param roles     placeholder に入れると確定した要素（id → role）
 * @param backgroundDim  Node が足す dim 矩形を先頭に置くか（§5.5）
 * @returns 図形名の列（backgroundDim なら要素数 + 1）
 */
export function assignNames(elements: Element[], roles: Map<string, RoleHint>, backgroundDim: boolean): string[] {
  const used = new Set<string>()
  const unique = (base: string): string => {
    let name = base
    for (let i = 2; used.has(name); i++) name = `${base}-${i}`
    used.add(name)
    return name
  }
  const out: string[] = []
  if (backgroundDim) out.push(unique(BACKGROUND_DIM))
  for (const e of elements) {
    const prop = sanitizeXmlText(e.name ?? '').trim()
    const role = roles.get(e.id)
    const base = prop || (role ? FIXED[role] : `${kindName(e)} ${elementNumber(e.id)}`)
    out.push(unique(base))
  }
  return out
}
