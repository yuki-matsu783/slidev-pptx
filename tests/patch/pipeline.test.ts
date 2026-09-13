// 後処理の列全体（native-export.md §3.1、§3.2）: 順番、何もしない Patch の往復、再圧縮、全部通したあとの OPC 検査
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { DOMParser } from '@xmldom/xmldom'
import { PATCHES, postProcess, xmlPatch } from '../../packages/slidev-addon-pptx/src/patch/index'
import type { Patch } from '../../packages/slidev-addon-pptx/src/patch/index'
import { check } from '../../packages/slidev-addon-pptx/src/opc/check'
import { openPptx, readFixturePptx } from '../helpers/pptx'
import { contextFor } from './helpers'

describe('patch/index: 列の定義', () => {
  it('順番は 1〜7 で固定', () => {
    expect(PATCHES.map((p) => p.name)).toEqual([
      'renameShapes',
      'dropEmptyPlaceholders',
      'dedupeParagraphProps',
      'applyAutofitScale',
      'splitNotesParagraphs',
      'replaceMaster',
      'rebuildContentTypes',
    ])
  })
})

describe('patch/index: postProcess', () => {
  it('何もしない Patch を通した結果は、全 XML が元と等価（§11 の 1: xmldom の往復）', async () => {
    // @xmldom/xmldom は空要素を <x/> に畳み、属性の改行を詰め、テキストの CRLF を LF に正規化する。
    // 「等価」= 要素名・属性の集合・テキスト（行末を LF に揃え、空白だけのノードは無視）が再帰的に同じ。
    // 元の文字列を DOM にしたものと、往復後の文字列を DOM にしたものを比べる（往復後同士を比べると冪等性しか見えない）
    const noop: Patch = xmlPatch('noop', /\.(xml|rels)$/, () => {})
    const before = await openPptx(readFixturePptx())
    const after = await openPptx(await postProcess(readFixturePptx(), [noop], contextFor()))
    expect(after.list().sort()).toEqual(before.list().sort())
    for (const path of before.list().filter((p) => /\.(xml|rels)$/.test(p))) {
      const diff = firstDifference(parse(await before.text(path)).documentElement!, parse(await after.text(path)).documentElement!)
      expect(diff, path).toBeUndefined()
    }
  })

  it('等価判定そのものの検査: 空要素の書き方と CRLF の違いは等価、属性やテキストの違いは非等価', () => {
    expect(firstDifference(parse('<a x="1"><b></b>t\r\n</a>').documentElement!, parse('<a x="1"><b/>t\n</a>').documentElement!)).toBeUndefined()
    expect(firstDifference(parse('<a x="1"/>').documentElement!, parse('<a x="2"/>').documentElement!)).toBeDefined()
    expect(firstDifference(parse('<a>t</a>').documentElement!, parse('<a>u</a>').documentElement!)).toBeDefined()
    expect(firstDifference(parse('<a><b/></a>').documentElement!, parse('<a/>').documentElement!)).toBeDefined()
  })

  it('往復は冪等（2 回通しても 1 回と同じバイト列）', async () => {
    const noop: Patch = xmlPatch('noop', /\.(xml|rels)$/, () => {})
    const once = await openPptx(await postProcess(readFixturePptx(), [noop], contextFor()))
    const twice = await openPptx(await postProcess(await postProcess(readFixturePptx(), [noop], contextFor()), [noop], contextFor()))
    for (const path of once.list().filter((p) => /\.(xml|rels)$/.test(p))) {
      expect(await twice.text(path), path).toBe(await once.text(path))
    }
  })

  it('バイナリ（media）は byte 単位で同じ', async () => {
    const before = await JSZip.loadAsync(readFixturePptx())
    const after = await JSZip.loadAsync(await postProcess(readFixturePptx(), [], contextFor()))
    const media = Object.keys(before.files).filter((f) => f.startsWith('ppt/media/'))
    expect(media.length).toBeGreaterThan(0)
    for (const m of media) {
      expect(Buffer.from(await after.file(m)!.async('uint8array'))).toEqual(Buffer.from(await before.file(m)!.async('uint8array')))
    }
  })

  it('再圧縮は DEFLATE（PptxGenJS の出力は STORE なので、何も変えなくても小さくなる）', async () => {
    const input = readFixturePptx()
    const out = await postProcess(input, [], contextFor())
    expect(out.length).toBeLessThan(input.length * 0.6)
    // 先頭のディレクトリ項目（_rels/ など。常に STORE）を飛ばし、最初のファイル項目の圧縮方式を見る
    expect(firstFileMethod(input)).toBe(0)
    expect(firstFileMethod(out)).toBe(8)
  })

  it('全部通したあとは OPC 検査の error が 0', async () => {
    const out = await postProcess(readFixturePptx(), PATCHES, contextFor())
    const results = await check(out)
    expect(results.filter((r) => r.level === 'error')).toEqual([])
  })

  it('Patch は列の順に走る（ctx に記録される順で確かめる）', async () => {
    const order: string[] = []
    const a: Patch = { name: 'a', run: () => { order.push('a') } }
    const b: Patch = { name: 'b', run: () => { order.push('b') } }
    await postProcess(readFixturePptx(), [a, b], contextFor())
    expect(order).toEqual(['a', 'b'])
  })
})

function parse(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml') as unknown as Document
}

/** 2 つの要素を再帰的に比べ、最初の違いを文字列で返す。同じなら undefined。
 *  空白だけのテキストノード・コメント・XML 宣言は見ない（PptxGenJS の出力には無い。`<a:t> </a:t>` を消す変換は検出できない） */
function firstDifference(a: Element, b: Element, path = a.nodeName): string | undefined {
  if (a.nodeName !== b.nodeName) return `${path}: name ${a.nodeName} != ${b.nodeName}`
  const attrs = (e: Element) => Object.fromEntries(Array.from(e.attributes).map((x) => [x.name, x.value]).sort())
  const aa = JSON.stringify(attrs(a))
  const ba = JSON.stringify(attrs(b))
  if (aa !== ba) return `${path}: attrs ${aa} != ${ba}`
  const kids = (e: Element) => Array.from(e.childNodes).filter((n) => n.nodeType === 1 || (n.nodeType === 3 && (n.nodeValue ?? '').trim() !== ''))
  const ak = kids(a)
  const bk = kids(b)
  if (ak.length !== bk.length) return `${path}: ${ak.length} children != ${bk.length}`
  for (let i = 0; i < ak.length; i++) {
    const x = ak[i]
    const y = bk[i]
    if (x.nodeType !== y.nodeType) return `${path}[${i}]: node type`
    if (x.nodeType === 3) {
      const norm = (s: string | null) => (s ?? '').replace(/\r\n?/g, '\n')
      if (norm(x.nodeValue) !== norm(y.nodeValue)) return `${path}[${i}]: text ${JSON.stringify(x.nodeValue)} != ${JSON.stringify(y.nodeValue)}`
    } else {
      const d = firstDifference(x as Element, y as Element, `${path}/${(x as Element).nodeName}[${i}]`)
      if (d) return d
    }
  }
  return undefined
}

/** ZIP のローカルヘッダを先頭から歩き、名前が / で終わらない最初の項目の圧縮方式（0 = STORE, 8 = DEFLATE）を返す */
function firstFileMethod(buf: Buffer): number {
  let off = 0
  while (off + 30 <= buf.length && buf.readUInt32LE(off) === 0x04034b50) {
    const method = buf.readUInt16LE(off + 8)
    const compressed = buf.readUInt32LE(off + 18)
    const nameLen = buf.readUInt16LE(off + 26)
    const extraLen = buf.readUInt16LE(off + 28)
    const name = buf.toString('utf8', off + 30, off + 30 + nameLen)
    if (!name.endsWith('/')) return method
    off += 30 + nameLen + extraLen + compressed
  }
  throw new Error('no file entry found')
}
