// 図形の名前（native-export.md §4.5）
import { describe, expect, it } from 'vitest'
import { assignNames } from '../../packages/slidev-addon-pptx/src/build/names'
import type { Element } from '../../packages/slidev-addon-pptx/src/types'

const el = (id: string, kind: Element['kind'], name = '', extra: Record<string, unknown> = {}): Element =>
  ({ id, name, kind, source: 'markdown', box: { x: 0, y: 0, w: 10, h: 10 }, boxSource: 'measured', ...extra }) as unknown as Element

describe('build/names: assignNames', () => {
  it('placeholder に入れると確定した枠は固定名 Title / Body / Body 2', () => {
    const els = [el('s1-e1', 'text'), el('s1-e2', 'text'), el('s1-e3', 'text')]
    const roles = new Map([['s1-e1', 'title'], ['s1-e2', 'body'], ['s1-e3', 'body2']] as const)
    expect(assignNames(els, roles, false)).toEqual(['Title', 'Body', 'Body 2'])
  })

  it('それ以外は <種類> <要素番号>（要素番号は id の M）', () => {
    const els = [el('s2-e3', 'text'), el('s2-e4', 'table'), el('s2-e5', 'image'), el('s2-e6', 'shape'), el('s2-e7', 'line'), el('s2-e9', 'image', '', { source: 'replaced', captureId: 's2-e9' })]
    expect(assignNames(els, new Map(), false)).toEqual(['Text 3', 'Table 4', 'Image 5', 'Shape 6', 'Line 7', 'Replaced 9'])
  })

  it('name prop があればそれ。同名は -2 -3 を足す', () => {
    const els = [el('s1-e1', 'shape', 'box'), el('s1-e2', 'shape', 'box'), el('s1-e3', 'shape', 'box')]
    expect(assignNames(els, new Map(), false)).toEqual(['box', 'box-2', 'box-3'])
  })

  it('Node が足す図形（Background dim）が列の先頭に入る', () => {
    const els = [el('s1-e1', 'text'), el('s1-e2', 'text')]
    const roles = new Map([['s1-e1', 'title']] as const)
    expect(assignNames(els, roles, true)).toEqual(['Background dim', 'Title', 'Text 2'])
  })

  it('name prop と固定名が衝突しても一意になる', () => {
    const els = [el('s1-e1', 'text'), el('s1-e2', 'text', 'Title')]
    const roles = new Map([['s1-e1', 'title']] as const)
    const names = assignNames(els, roles, false)
    expect(names[0]).toBe('Title')
    expect(new Set(names).size).toBe(2)
  })

  it('XML に入れられない文字は name から落ちる', () => {
    const els = [el('s1-e1', 'text', `na${String.fromCharCode(1)}me`)]
    expect(assignNames(els, new Map(), false)).toEqual(['name'])
  })
})
