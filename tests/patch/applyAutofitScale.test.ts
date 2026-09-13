// 後処理 4: applyAutofitScale（native-export.md §3.2、§3.3）
import { describe, expect, it } from 'vitest'
import { renameShapes } from '../../packages/slidev-addon-pptx/src/patch/patches/renameShapes'
import { applyAutofitScale } from '../../packages/slidev-addon-pptx/src/patch/patches/applyAutofitScale'
import { openPptx, shapeNamed, els } from '../helpers/pptx'
import { contextFor, runPatches } from './helpers'

describe('patch/applyAutofitScale', () => {
  it('ctx.autofit の図形名で引き、既存の <a:normAutofit> に fontScale と lnSpcReduction を足す', async () => {
    const ctx = contextFor({ autofit: { 1: { 'Text 2': { fontScale: 85000, lnSpcReduction: 10000 } } } })
    const p = await openPptx(await runPatches([renameShapes, applyAutofitScale], ctx))
    const sp = shapeNamed(await p.xml('ppt/slides/slide1.xml'), 'Text 2')!
    const bodyPr = els(sp, 'a', 'bodyPr')[0]
    const autofits = Array.from(bodyPr.childNodes).filter((n) => n.nodeType === 1 && /Autofit$/i.test((n as Element).localName))
    expect(autofits).toHaveLength(1)
    const na = els(bodyPr, 'a', 'normAutofit')[0]
    expect(na.getAttribute('fontScale')).toBe('85000')
    expect(na.getAttribute('lnSpcReduction')).toBe('10000')
  })

  it('normAutofit が無い枠には 1 つ作る（2 つにはしない）', async () => {
    // slide2 の自由配置テキストは fit 指定なし → normAutofit が無い
    const ctx = contextFor({ autofit: { 2: { 'Text 1': { fontScale: 50000, lnSpcReduction: 20000 } } } })
    const p = await openPptx(await runPatches([renameShapes, applyAutofitScale], ctx))
    const sp = shapeNamed(await p.xml('ppt/slides/slide2.xml'), 'Text 1')!
    const bodyPr = els(sp, 'a', 'bodyPr')[0]
    expect(els(bodyPr, 'a', 'normAutofit')).toHaveLength(1)
    expect(els(bodyPr, 'a', 'normAutofit')[0].getAttribute('fontScale')).toBe('50000')
  })

  it('autofit に無い図形は触らない', async () => {
    const p = await openPptx(await runPatches([renameShapes, applyAutofitScale], contextFor({ autofit: {} })))
    const sp = shapeNamed(await p.xml('ppt/slides/slide1.xml'), 'Text 2')!
    const na = els(sp, 'a', 'normAutofit')[0]
    expect(na.hasAttribute('fontScale')).toBe(false)
  })
})
