// MasterSwap の型（native-export.md §5.4）。vitest の typecheck（tests/vitest.config.ts の typecheck.include）で検査する。
// 実行時の expectTypeOf は no-op なので、型の主張は *.test-d.ts に置く
import { describe, expectTypeOf, it } from 'vitest'
import type { MasterSwap } from '../../packages/slidev-addon-pptx/src/patch/patches/replaceMaster'

describe('MasterSwap', () => {
  it('template / layoutMap / placeholderMap を持つ', () => {
    expectTypeOf<MasterSwap>().toHaveProperty('template')
    expectTypeOf<MasterSwap['template']>().toEqualTypeOf<Buffer>()
    expectTypeOf<MasterSwap['layoutMap']>().toEqualTypeOf<Record<string, string>>()
    expectTypeOf<MasterSwap['placeholderMap'][string]>().toEqualTypeOf<Record<'title' | 'body' | 'body2', { idx: number; type: string }>>()
  })
})
