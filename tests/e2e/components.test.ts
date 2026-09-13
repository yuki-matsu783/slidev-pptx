// PPT 部品（ppt-components.md §1）を使うデッキ。アドオンの実装（フェーズ 4）が入るまで skip。
// 有効化するときは `describe.skip` を `describe` に戻す。
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import type { Browser } from 'playwright-chromium'
import { collect } from '../../packages/slidev-addon-pptx/src/collect/index'
import type { Capture, TextElement, ShapeElement, ImageElement, TableElement, LineElement } from '../../packages/slidev-addon-pptx/src/types'
import { COMPONENTS, launch, openPrint, startSlidev } from './helpers'
import type { Running } from './helpers'

let server: Running
let browser: Browser
let capture: Capture

describe.skip('collect: PPT 部品（フェーズ 4 で有効化）', () => {
  beforeAll(async () => {
    server = await startSlidev(COMPONENTS)
    browser = await launch()
    const page = await openPrint(browser, server.base)
    capture = await page.evaluate(collect)
  }, 120_000)
  afterAll(async () => {
    await browser?.close()
    await server?.close()
  })

  const els = () => capture.slides[0].elements

  it('PptText: data-ppt-box の座標指定は測らずにそのまま、h は実測。fill / radius / padding が frame に', () => {
    const lead = els().find((e) => e.name === 'lead') as TextElement
    expect(lead.kind).toBe('text')
    expect(lead.source).toBe('ppt')
    expect(lead.boxSource).toBe('prop')
    expect(lead.box).toMatchObject({ x: 60, y: 120, w: 400 })
    expect(lead.frame.fill?.color).toBe('#eeeeff')
    expect(lead.frame.radius).toBe(8)
    expect(lead.frame.inset).toEqual([12, 12, 12, 12])
    expect(lead.paragraphs[0].runs[0].bold).toBe(true)
  })
  it('PptShape: type / fill / line none、中の文字', () => {
    const arrow = els().find((e) => e.name === 'arrow') as ShapeElement
    expect(arrow.kind).toBe('shape')
    expect(arrow.shape).toBe('rightArrow')
    expect(arrow.frame.fill?.color).toBe('#3b82f6')
    expect(arrow.frame.line).toBeUndefined()
    expect(arrow.paragraphs?.[0].runs[0].text).toBe('次へ')
  })
  it('PptShape type="line" は line 要素で、tail が付く', () => {
    const line = els().find((e) => e.kind === 'line') as LineElement
    expect(line.from).toEqual({ x: 60, y: 200 })
    expect(line.to).toEqual({ x: 460, y: 200 })
    expect(line.line.tail).toBe('arrow')
  })
  it('PptImage: src は絶対 URL、fit と alt', () => {
    const img = els().find((e) => e.kind === 'image' && (e as ImageElement).alt === '背景') as ImageElement
    expect(img.src).toMatch(/^http.*\/bg\.png$/)
    expect(img.fit).toBe('cover')
  })
  it('PptTable: rows から表、1 行目が見出し', () => {
    const tbl = els().find((e) => e.kind === 'table') as TableElement
    expect(tbl.headerRows).toBe(1)
    expect(tbl.rows.map((r) => r.map((c) => c.paragraphs[0].runs[0].text))).toEqual([['名前', '用途'], ['default', '通常'], ['cover', '表紙']])
  })
  it('export="image" の部品は画像への置き換え（explicit）', () => {
    expect(els().some((e) => e.kind === 'image' && (e as ImageElement).reason === 'explicit')).toBe(true)
  })
  it('部品の入れ子は W-NESTED-PPT、PptText の中の表は W-PPT-CONTENT', () => {
    const codes = capture.slides[0].warnings.map((w) => w.code)
    expect(codes).toContain('W-NESTED-PPT')
    expect(codes).toContain('W-PPT-CONTENT')
  })
})
