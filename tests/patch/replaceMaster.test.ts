// 後処理 6: replaceMaster は今回 no-op（native-export.md §5.4）。
// no-op でも MasterSwap の型と、rels を Type で引く補助関数は動くことを固定する。
import { describe, expect, it } from 'vitest'
import { DOMParser } from '@xmldom/xmldom'
import { replaceMaster } from '../../packages/slidev-addon-pptx/src/patch/patches/replaceMaster'
import type { MasterSwap } from '../../packages/slidev-addon-pptx/src/patch/patches/replaceMaster'
import { findRelByType } from '../../packages/slidev-addon-pptx/src/patch/zip'
import { openPptx, readFixturePptx } from '../helpers/pptx'
import { runPatches } from './helpers'

const REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const RT = (n: string) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${n}`

describe('patch/replaceMaster (no-op)', () => {
  it('MasterSwap の型（型の検査は replaceMaster.test-d.ts。ここでは値が作れることだけ）', () => {
    const swap: MasterSwap = {
      template: Buffer.alloc(0),
      layoutMap: { cover: 'Title Slide' },
      placeholderMap: { cover: { title: { idx: 0, type: 'ctrTitle' }, body: { idx: 1, type: 'subTitle' }, body2: { idx: 2, type: 'body' } } },
    }
    expect(Object.keys(swap).sort()).toEqual(['layoutMap', 'placeholderMap', 'template'])
  })

  it('swap を渡さなければ ZIP の全パートが元のまま', async () => {
    const before = await openPptx(readFixturePptx())
    const after = await openPptx(await runPatches([replaceMaster]))
    expect(after.list().sort()).toEqual(before.list().sort())
    for (const path of before.list().filter((p) => p.endsWith('.xml') || p.endsWith('.rels'))) {
      expect(await after.text(path)).toBe(await before.text(path))
    }
  })
})

describe('patch/zip: findRelByType', () => {
  const rels = new DOMParser().parseFromString(
    `<?xml version="1.0"?><Relationships xmlns="${REL}">` +
      `<Relationship Id="rId1" Type="${RT('hyperlink')}" Target="https://sli.dev" TargetMode="External"/>` +
      `<Relationship Id="rId2" Type="${RT('slideLayout')}" Target="../slideLayouts/slideLayout4.xml"/>` +
      `<Relationship Id="rId3" Type="${RT('notesSlide')}" Target="../notesSlides/notesSlide1.xml"/>` +
      `</Relationships>`,
    'application/xml',
  ) as unknown as Document

  it('Type で引く（rId は決め打ちにしない）', () => {
    const rel = findRelByType(rels, 'slideLayout')
    expect(rel?.getAttribute('Id')).toBe('rId2')
    expect(rel?.getAttribute('Target')).toBe('../slideLayouts/slideLayout4.xml')
  })

  it('無ければ undefined', () => {
    expect(findRelByType(rels, 'image')).toBeUndefined()
  })

  it('完全な URI でも短い名前でも引ける', () => {
    expect(findRelByType(rels, RT('notesSlide'))?.getAttribute('Id')).toBe('rId3')
  })
})
