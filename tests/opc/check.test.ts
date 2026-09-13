// OPC 整合チェッカ（native-export.md §3.4）の規則 C1〜C14 を 1 規則ずつ、
// 「通る / 落ちる」の対で検査する。ZIP は jszip で最小に組み立てる。
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { check } from '../../packages/slidev-addon-pptx/src/opc/check'
import type { CheckResult } from '../../packages/slidev-addon-pptx/src/types'
import { readFixturePptx } from '../helpers/pptx'

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
const T = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  pres: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
}
const RT = (n: string) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${n}`

interface MinimalOpts {
  slideXml?: string
  slideRels?: string
  layoutXml?: string
  masterXml?: string
  masterRels?: string
  presXml?: string
  presRels?: string
  contentTypes?: string
  extraParts?: Record<string, string | Uint8Array>
}

function sp(id: number, name: string, inner = '', ph = ''): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>` +
    `<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr>${inner || '<a:p><a:r><a:rPr lang="ja-JP"/><a:t>x</a:t></a:r></a:p>'}</p:txBody></p:sp>`
}

function slideXml(shapes: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`
}

/** 整合の取れた最小の PPTX（スライド 1 枚、レイアウト 1 枚、マスター 1 枚）。引数で 1 か所だけ壊す */
async function minimalPptx(o: MinimalOpts = {}): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', o.contentTypes ?? `<?xml version="1.0"?><Types xmlns="${CT}">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="${T.pres}"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${T.master}"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${T.layout}"/>` +
    `<Override PartName="/ppt/slides/slide1.xml" ContentType="${T.slide}"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="${T.theme}"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${RT('officeDocument')}" Target="ppt/presentation.xml"/></Relationships>`)
  zip.file('ppt/presentation.xml', o.presXml ?? `<?xml version="1.0"?><p:presentation xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', o.presRels ?? `<?xml version="1.0"?><Relationships xmlns="${REL}">` +
    `<Relationship Id="rId1" Type="${RT('slideMaster')}" Target="slideMasters/slideMaster1.xml"/>` +
    `<Relationship Id="rId2" Type="${RT('slide')}" Target="slides/slide1.xml"/>` +
    `<Relationship Id="rId3" Type="${RT('theme')}" Target="theme/theme1.xml"/></Relationships>`)
  zip.file('ppt/slideMasters/slideMaster1.xml', o.masterXml ?? `<?xml version="1.0"?><p:sldMaster xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}">` +
    `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`)
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', o.masterRels ?? `<?xml version="1.0"?><Relationships xmlns="${REL}">` +
    `<Relationship Id="rId1" Type="${RT('slideLayout')}" Target="../slideLayouts/slideLayout1.xml"/>` +
    `<Relationship Id="rId2" Type="${RT('theme')}" Target="../theme/theme1.xml"/></Relationships>`)
  zip.file('ppt/slideLayouts/slideLayout1.xml', o.layoutXml ?? `<?xml version="1.0"?><p:sldLayout xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld name="default"><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    sp(2, 'title', '', '<p:ph type="title" idx="100"/>') + sp(3, 'body', '', '<p:ph type="body" idx="101"/>') +
    `</p:spTree></p:cSld></p:sldLayout>`)
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${RT('slideMaster')}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`)
  zip.file('ppt/slides/slide1.xml', o.slideXml ?? slideXml(sp(2, 'Title', '', '<p:ph type="title" idx="100"/>') + sp(3, 'Text 3')))
  zip.file('ppt/slides/_rels/slide1.xml.rels', o.slideRels ?? `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${RT('slideLayout')}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`)
  zip.file('ppt/theme/theme1.xml', `<?xml version="1.0"?><a:theme xmlns:a="${A}" name="t"><a:themeElements/></a:theme>`)
  for (const [p, c] of Object.entries(o.extraParts ?? {})) zip.file(p, c)
  return zip.generateAsync({ type: 'nodebuffer' })
}

const rulesOf = (r: CheckResult[]) => r.map((x) => x.rule)
const errorsOf = (r: CheckResult[]) => r.filter((x) => x.level === 'error').map((x) => x.rule)

describe('opc/check: 整合の取れた最小 PPTX', () => {
  it('error も warn も出ない', async () => {
    expect(await check(await minimalPptx())).toEqual([])
  })
  it('結果は { rule, part, message, level } の形', async () => {
    const r = await check(await minimalPptx({ slideXml: slideXml(sp(2, 'a') + sp(2, 'b')) }))
    expect(r[0]).toMatchObject({ rule: 'C6', part: 'ppt/slides/slide1.xml', level: 'error' })
    expect(typeof r[0].message).toBe('string')
  })
})

describe('opc/check: 規則ごと', () => {
  it('C1: parse できない XML（@xmldom/xmldom は既定では例外を投げず document を返すので、errorHandler / onError で拾う前提）', async () => {
    const r = await check(await minimalPptx({ slideXml: '<p:sld><p:cSld>' }))
    expect(errorsOf(r)).toContain('C1')
  })

  it('C2: Override が実在しないパートを指す', async () => {
    const r = await check(await minimalPptx({ contentTypes: undefined, extraParts: {} }).then(async (buf) => {
      const zip = await JSZip.loadAsync(buf)
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      zip.file('[Content_Types].xml', ct.replace('</Types>', `<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="${T.master}"/></Types>`))
      return zip.generateAsync({ type: 'nodebuffer' })
    }))
    expect(errorsOf(r)).toContain('C2')
  })

  it('C3: Default にも Override にも無いパート', async () => {
    const r = await check(await minimalPptx({ extraParts: { 'ppt/media/image1.webp': new Uint8Array([0]) } }))
    expect(errorsOf(r)).toContain('C3')
  })

  it('C4: rels の Target が実在しない（External は対象外）', async () => {
    const bad = await check(await minimalPptx({ slideRels: `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${RT('slideLayout')}" Target="../slideLayouts/slideLayout9.xml"/></Relationships>` }))
    expect(errorsOf(bad)).toContain('C4')
    const ok = await check(await minimalPptx({ slideRels: `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${RT('slideLayout')}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${RT('hyperlink')}" Target="https://sli.dev" TargetMode="External"/></Relationships>` }))
    expect(errorsOf(ok)).not.toContain('C4')
  })

  it('C5: スライドに slideLayout の関係が無い / presentation に slideMaster が無い', async () => {
    const r1 = await check(await minimalPptx({ slideRels: `<?xml version="1.0"?><Relationships xmlns="${REL}"/>` }))
    expect(errorsOf(r1)).toContain('C5')
    const r2 = await check(await minimalPptx({ presRels: `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId2" Type="${RT('slide')}" Target="slides/slide1.xml"/></Relationships>` }))
    expect(errorsOf(r2)).toContain('C5')
  })

  it('C6: cNvPr id の重複と 0', async () => {
    const dup = await check(await minimalPptx({ slideXml: slideXml(sp(2, 'a') + sp(2, 'b')) }))
    expect(errorsOf(dup)).toContain('C6')
    const zero = await check(await minimalPptx({ slideXml: slideXml(sp(0, 'a')) }))
    expect(errorsOf(zero)).toContain('C6')
  })

  it('C6: レイアウト・マスター・ノートのパートも見る', async () => {
    const r = await check(await minimalPptx({ layoutXml: `<?xml version="1.0"?><p:sldLayout xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld name="x"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${sp(5, 'a')}${sp(5, 'b')}</p:spTree></p:cSld></p:sldLayout>` }))
    expect(r.find((x) => x.rule === 'C6')?.part).toBe('ppt/slideLayouts/slideLayout1.xml')
  })

  it('C7a: presentation.xml の sldId / sldMasterId の r:id が rels に無い', async () => {
    const r = await check(await minimalPptx({ presXml: `<?xml version="1.0"?><p:presentation xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId9"/></p:sldIdLst></p:presentation>` }))
    expect(errorsOf(r)).toContain('C7a')
  })

  it('C7b: slideMasterN.xml の sldLayoutId の r:id が rels に無い', async () => {
    const r = await check(await minimalPptx({ masterRels: `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId7" Type="${RT('slideLayout')}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>` }))
    expect(errorsOf(r)).toContain('C7b')
  })

  it('C8 (warn): スライドの ph idx がレイアウトに無い。idx 無しの type="title" は可', async () => {
    const warn = await check(await minimalPptx({ slideXml: slideXml(sp(2, 'Body', '', '<p:ph type="body" idx="999"/>')) }))
    expect(warn.find((x) => x.rule === 'C8')?.level).toBe('warn')
    const ok = await check(await minimalPptx({ slideXml: slideXml(sp(2, 'Title', '', '<p:ph type="title"/>')) }))
    expect(rulesOf(ok)).not.toContain('C8')
  })

  it('C9 (warn): cx=0 cy=0 の図形', async () => {
    const x = sp(2, 'z').replace('cx="914400" cy="914400"', 'cx="0" cy="0"')
    const r = await check(await minimalPptx({ slideXml: slideXml(x) }))
    expect(r.find((v) => v.rule === 'C9')?.level).toBe('warn')
  })

  it('C9: spTree 自身の <p:grpSpPr><a:xfrm> の cx=0 cy=0 は対象外（PptxGenJS は常に 0 で出す）', async () => {
    const withGrp = slideXml(sp(2, 'a')).replace('<p:grpSpPr/>', '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>')
    const r = await check(await minimalPptx({ slideXml: withGrp }))
    expect(rulesOf(r)).not.toContain('C9')
  })

  it('C10: r: 名前空間の属性値がそのパートの rels に無い（r:id / r:embed / r:link）', async () => {
    const withLink = sp(2, 'L', `<a:p><a:r><a:rPr lang="ja-JP"><a:hlinkClick r:id="rId9"/></a:rPr><a:t>x</a:t></a:r></a:p>`)
    const r = await check(await minimalPptx({ slideXml: slideXml(withLink) }))
    expect(errorsOf(r)).toContain('C10')
    const pic = `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId5"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr></p:pic>`
    const r2 = await check(await minimalPptx({ slideXml: slideXml(sp(2, 'a') + pic) }))
    expect(errorsOf(r2)).toContain('C10')
  })

  it('C11: bodyPr に autofit が 2 つ', async () => {
    const x = sp(2, 'a').replace('<a:bodyPr><a:normAutofit/></a:bodyPr>', '<a:bodyPr><a:normAutofit/><a:spAutoFit/></a:bodyPr>')
    const r = await check(await minimalPptx({ slideXml: slideXml(x) }))
    expect(errorsOf(r)).toContain('C11')
  })

  it('C12: 1 スライド内で ph idx が重複（idx 無しは対象外）', async () => {
    const dup = await check(await minimalPptx({ slideXml: slideXml(sp(2, 'a', '', '<p:ph type="title" idx="100"/>') + sp(3, 'b', '', '<p:ph type="body" idx="100"/>')) }))
    expect(errorsOf(dup)).toContain('C12')
    const noIdx = await check(await minimalPptx({ slideXml: slideXml(sp(2, 'a', '', '<p:ph type="title"/>') + sp(3, 'b', '', '<p:ph type="title"/>')) }))
    expect(errorsOf(noIdx)).not.toContain('C12')
  })

  it('C13: テキストノードと属性値の制御文字', async () => {
    const inText = sp(2, 'a', `<a:p><a:r><a:rPr lang="ja-JP"/><a:t>badchar</a:t></a:r></a:p>`)
    expect(errorsOf(await check(await minimalPptx({ slideXml: slideXml(inText) })))).toContain('C13')
    const inAttr = sp(2, 'name')
    expect(errorsOf(await check(await minimalPptx({ slideXml: slideXml(inAttr) })))).toContain('C13')
  })

  it('C14: <a:p> 直下の <a:pPr> は高々 1 つで最初の子', async () => {
    const two = sp(2, 'a', `<a:p><a:pPr/><a:r><a:rPr lang="ja-JP"/><a:t>x</a:t></a:r><a:pPr/></a:p>`)
    expect(errorsOf(await check(await minimalPptx({ slideXml: slideXml(two) })))).toContain('C14')
    const late = sp(2, 'a', `<a:p><a:r><a:rPr lang="ja-JP"/><a:t>x</a:t></a:r><a:pPr/></a:p>`)
    expect(errorsOf(await check(await minimalPptx({ slideXml: slideXml(late) })))).toContain('C14')
  })
})

describe('opc/check: PptxGenJS の素の出力', () => {
  it('fixture の sample.pptx は後処理前なので C2 / C6 / C14 で落ちる', async () => {
    const r = await check(readFixturePptx())
    const e = errorsOf(r)
    expect(e).toContain('C2')
    expect(e).toContain('C6')
    expect(e).toContain('C14')
  })
})
