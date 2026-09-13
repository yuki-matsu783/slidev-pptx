// Capture → PptxGenJS（native-export.md §4、§5、§6、§8）。単位の換算はここより下（units.ts）だけがやる。
import PptxGenJS from 'pptxgenjs'
import type {
  Box, Canvas, Capture, Cell, DeckData, Element, ImageElement, LineElement, Paragraph, Report, RoleHint, Run, ShapeElement, SlideCapture,
  TableElement, TextElement, Warning,
} from '../types.ts'
import { emptyReport } from '../types.ts'
import type { PatchContext, AutofitScale } from '../patch/index.ts'
import { emu, pt, inch, clampBox, textMargin, cellMargin, toDashType } from './units.ts'
import { fontFor, themeFonts } from './fonts.ts'
import { assignNames } from './names.ts'
import { sanitizeXmlText } from './sanitize.ts'
import { notesText } from './notes.ts'
import { resolveLayout } from './layout.ts'
import { defineMasters, masterFor, hasPlaceholder, placeholderBox } from './masters.ts'
import type { LayoutName } from './masters.ts'

export interface Asset {
  data: string
  width?: number
  height?: number
  selector?: string
}

export interface BuildOptions {
  /** captureId / src / `background:<no>` → data URL（または撮影の情報つき） */
  assets: Record<string, string | Asset>
  /** run の lang の既定（スライドに lang が無いとき） */
  lang?: string
  /** options.utils.getLayouts() のキー一覧 */
  layouts: string[]
  /** Report の output / 版 */
  output?: string
  slidevVersion?: string
}

export interface BuildResult {
  pptx: PptxGenJS
  ctx: PatchContext
}

const OFFSLIDE_WARN_PX = 5
const DIM_TRANSPARENCY = 55

export async function build(input: Capture, data: DeckData, opts: BuildOptions): Promise<BuildResult> {
  // 入力を壊さない（リンクの番号の変換などで書き換えるため複製を取る）
  const capture: Capture = structuredClone(input)
  const canvas = capture.canvas
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.theme = themeFonts()
  if (data.config?.title) pptx.title = sanitizeXmlText(data.config.title)
  if (data.config?.author) pptx.author = sanitizeXmlText(data.config.author)
  defineMasters(pptx, canvas.width)

  const report: Report = emptyReport(opts.output ?? '')
  report.slidev = opts.slidevVersion ?? ''
  report.pptxgenjs = pptxVersion()
  const ctx: PatchContext = { capture, shapeNames: {}, autofit: {}, report }
  const drop = (key: string) => {
    report.dropped[key] = (report.dropped[key] ?? 0) + 1
  }
  const warn = (slide: number, w: Warning, name?: string) => {
    report.warnings.push({ ...w, slide, name })
  }

  const transition = data.config?.transition ?? data.slides.find((s) => s.frontmatter?.transition)?.frontmatter?.transition
  if (transition) {
    warn(0, { code: 'W-TRANSITION', message: `transition "${transition}" は PPTX に出せない（PptxGenJS に API が無い）` })
    drop('transition')
  }
  if (data.config?.colorSchema === 'dark') warn(0, { code: 'W-DARK', message: 'colorSchema: dark のデッキ。測った色のまま出す' })

  const slides = [...capture.slides].sort((a, b) => a.no - b.no)
  report.slides = slides.length
  let keptTotal = 0
  // PptxGenJS の hyperlink.slide は PPTX の中での順番（slideN.xml）。range で絞ると元の番号とずれるので写像する。
  // 後処理（shapeNames / autofit）も PPTX の番号で引く。Report の warnings / replacements は元の番号のまま（slideMap で突き合わせる）
  const slideIndexByNo = new Map(slides.map((s, i) => [s.no, i + 1]))
  for (const [no, idx] of slideIndexByNo) report.slideMap[idx] = no
  const mapLink = (link: Run['link'] | undefined, no: number, id: string): Run['link'] | undefined => {
    if (!link || !('slide' in link)) return link
    const target = slideIndexByNo.get(link.slide)
    if (target) return { slide: target }
    warn(no, { code: 'W-LINK', elementId: id, message: `スライド ${link.slide} へのリンクは書き出しの範囲外なので外した` })
    return undefined
  }

  for (const sc of slides) {
    const index = sc.no - 1
    const layoutName = resolveLayout(index, data, opts.layouts)
    const master = masterFor(layoutName)
    if (master === 'blank' && layoutName !== 'blank') {
      warn(sc.no, { code: 'W-LAYOUT', message: `レイアウト "${layoutName}" は対応表に無いので blank にした` })
    }
    const slide = pptx.addSlide({ masterName: master })
    const pptxNo = slideIndexByNo.get(sc.no)!
    const lang = sc.lang || opts.lang || 'ja-JP'
    if (sc.zoom && sc.zoom !== 1) report.zoom[sc.no] = sc.zoom
    for (const [k, v] of Object.entries(sc.dropped ?? {})) report.dropped[k] = (report.dropped[k] ?? 0) + v

    // 背景（§5.5）
    const fmBg = data.slides[index]?.frontmatter?.background
    let backgroundDim = false
    if (typeof fmBg === 'string' && fmBg) {
      if (/^(#|rgb|hsl)/.test(fmBg)) {
        slide.background = { color: hex(fmBg) }
      } else {
        const asset = assetOf(opts.assets, `background:${sc.no}`) ?? (fmBg.startsWith('data:') ? { data: fmBg } : undefined)
        if (asset) {
          slide.background = { data: asset.data }
          backgroundDim = true
          drop('background-crop')
        } else {
          warn(sc.no, { code: 'W-IMAGE', message: `背景画像を取得できない: ${fmBg}` })
          slide.background = { color: hex(sc.backgroundColor) }
        }
      }
    } else {
      slide.background = { color: hex(sc.backgroundColor) }
    }

    // 要素の前処理: 自由配置の寄せ・落とし、role の確定
    const roles = new Map<string, RoleHint>()
    const usedRoles = new Set<RoleHint>()
    const boxes = new Map<string, Box>()
    const kept: Element[] = []
    for (const e of sc.elements) {
      let placeholder = false
      if (e.kind === 'text' && e.roleHint) {
        if (!hasPlaceholder(master, e.roleHint)) {
          if (master !== 'blank') warn(sc.no, { code: 'W-LAYOUT', elementId: e.id, message: `レイアウト "${master}" に placeholder "${e.roleHint}" が無いので自由配置にした` })
        } else if (usedRoles.has(e.roleHint)) {
          // 同じ role を 2 つ入れると <p:ph idx> が重複する（C12）。2 つ目以降は自由配置
          warn(sc.no, { code: 'W-LAYOUT', elementId: e.id, message: `placeholder "${e.roleHint}" は既に使われているので自由配置にした` })
        } else {
          roles.set(e.id, e.roleHint)
          usedRoles.add(e.roleHint)
          placeholder = true
        }
      }
      if (placeholder) {
        kept.push(e)
        continue
      }
      if (e.kind === 'line') {
        // 線は端点で寄せる（負の EMU はインチ扱いになって壊れる。§4.1）
        const x1 = Math.min(e.from.x, e.to.x)
        const y1 = Math.min(e.from.y, e.to.y)
        const x2 = Math.max(e.from.x, e.to.x)
        const y2 = Math.max(e.from.y, e.to.y)
        const cx1 = Math.max(0, Math.min(canvas.width, x1))
        const cx2 = Math.max(0, Math.min(canvas.width, x2))
        const cy1 = Math.max(0, Math.min(canvas.height, y1))
        const cy2 = Math.max(0, Math.min(canvas.height, y2))
        const shift = Math.max(cx1 - x1, x2 - cx2, cy1 - y1, y2 - cy2)
        // どちらかの軸で全体がスライドの外なら落とす（水平線が y<0、縦線が x>幅 など）
        if (x2 <= 0 || x1 >= canvas.width || y2 < 0 || y1 > canvas.height || (cx2 - cx1 <= 0 && cy2 - cy1 <= 0)) {
          warn(sc.no, { code: 'W-HIDDEN', elementId: e.id, message: 'スライドの外にあるので飛ばした' })
          continue
        }
        if (shift >= OFFSLIDE_WARN_PX) warn(sc.no, { code: 'W-OFFSLIDE', elementId: e.id, message: `スライドの外に ${Math.round(shift)} px 掛かるので内側に寄せた` })
        const flip = (e.from.x < e.to.x) !== (e.from.y < e.to.y) && e.from.y !== e.to.y && e.from.x !== e.to.x
        e.from = { x: cx1, y: flip ? cy2 : cy1 }
        e.to = { x: cx2, y: flip ? cy1 : cy2 }
        kept.push(e)
        continue
      }
      const c = clampBox(e.box, canvas)
      if (c.dropped) {
        warn(sc.no, { code: 'W-HIDDEN', elementId: e.id, message: 'スライドの外にあるので飛ばした' })
        continue
      }
      if (c.shift >= OFFSLIDE_WARN_PX) {
        warn(sc.no, { code: 'W-OFFSLIDE', elementId: e.id, message: `スライドの外に ${Math.round(c.shift)} px 掛かるので内側に寄せた` })
      }
      boxes.set(e.id, c.box)
      kept.push(e)
    }

    const names = assignNames(kept, roles, backgroundDim)
    ctx.shapeNames[pptxNo] = names
    let ni = 0
    if (backgroundDim) {
      slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: '100%',
        h: '100%',
        fill: { color: '000000', transparency: DIM_TRANSPARENCY },
        line: { color: '000000', width: 0, transparency: 100 },
        objectName: names[ni],
      })
      ni++
    }

    const nameOf = new Map<string, string>()
    for (const e of kept) nameOf.set(e.id, names[ni++])

    // スライドへのリンクの番号を PPTX の順番に変換する（範囲外は外す）
    for (const e of kept) {
      e.link = mapLink(e.link, sc.no, e.id)
      const paragraphs = e.kind === 'text' || e.kind === 'shape' ? e.paragraphs : e.kind === 'table' ? e.rows.flat().flatMap((c) => c.paragraphs) : undefined
      for (const p of paragraphs ?? []) for (const r of p.runs) r.link = mapLink(r.link, sc.no, e.id)
    }

    for (const e of kept) {
      const name = nameOf.get(e.id)!
      const box = boxes.get(e.id) ?? e.box
      switch (e.kind) {
        case 'text':
          addText(pptx, slide, e, box, roles.get(e.id), master, name, lang, canvas, sc, pptxNo, ctx)
          break
        case 'shape':
          addShape(pptx, slide, e, box, name, lang, canvas)
          break
        case 'line':
          addLine(pptx, slide, e, name, canvas)
          break
        case 'image':
          addImage(pptx, slide, e, box, name, canvas, sc, opts, report)
          break
        case 'table':
          addTable(slide, e, box, name, lang, canvas)
          break
      }
    }

    for (const w of sc.warnings) warn(sc.no, w, w.elementId ? nameOf.get(w.elementId) : undefined)

    const note = data.slides[index]?.note
    if (note) slide.addNotes(sanitizeXmlText(notesText(note)))
    keptTotal += kept.length
  }

  // native = 出せた要素（落としたものは数えない）− 画像への置き換え
  report.replaced = report.replacements.length
  report.native = keptTotal - report.replaced
  return { pptx, ctx }
}

// ---------------------------------------------------------------- テキスト

function addText(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  e: TextElement,
  box: Box,
  role: RoleHint | undefined,
  master: LayoutName,
  name: string,
  lang: string,
  canvas: Canvas,
  sc: SlideCapture,
  pptxNo: number,
  ctx: PatchContext,
): void {
  const runs = toRuns(e.paragraphs, lang, canvas, e.link)
  const o: PptxGenJS.TextPropsOptions = {
    fit: 'shrink',
    valign: e.valign,
    margin: textMargin(e.frame.inset, canvas),
    lang,
  }
  if (role) {
    o.placeholder = role
  } else {
    Object.assign(o, boxProps(box, canvas))
    o.isTextBox = true
  }
  applyFrame(pptx, o, e.frame, canvas)
  if (e.rotate) o.rotate = e.rotate
  slide.addText(runs.length ? runs : [{ text: '' }], o)

  // 縮小率（§3.3）。2 条件の大きいほうの不足率を採る:
  //  A. Slidev 側で既に溢れている（fit.contentHeight > fit.boxHeight）
  //  B. 出す枠（placeholder は対応表の h。対応表の px はキャンバス px と同じ空間。自由配置は寄せた枠の h）より内容が高い
  // フォントの違い（計測環境と PowerPoint）で数 % は動くので、5% を超えて溢れるときだけ書く
  const ph = role ? placeholderBox(master, role) : undefined
  const targetH = ph ? ph.h : box.h
  const contentH = e.fit?.contentHeight ?? 0
  const ratios: number[] = []
  if (contentH > 0 && e.fit.boxHeight > 0 && contentH > e.fit.boxHeight) ratios.push(e.fit.boxHeight / contentH)
  if (contentH > 0 && targetH > 0 && contentH > targetH) ratios.push(targetH / contentH)
  const ratio = Math.min(...ratios, 1)
  if (ratio < 0.95) {
    const scale: AutofitScale = {
      fontScale: Math.max(25000, Math.floor(ratio * 100) * 1000),
      lnSpcReduction: 0,
    }
    scale.lnSpcReduction = scale.fontScale < 90000 ? 20000 : 10000
    ;(ctx.autofit[pptxNo] ??= {})[name] = scale
    ctx.report.warnings.push({ code: 'W-OVERFLOW', elementId: e.id, slide: sc.no, name, message: `枠に収まらないので縮小率 ${scale.fontScale / 1000}% を書いた` })
  }
}

function toRuns(paragraphs: Paragraph[], lang: string, canvas: Canvas, elementLink?: Run['link']): PptxGenJS.TextProps[] {
  const out: PptxGenJS.TextProps[] = []
  paragraphs.forEach((p, pi) => {
    const runs = p.runs.length ? p.runs : [{ text: '', size: 17.6, color: '#000000', bold: false, italic: false, underline: false, strike: false, code: false } as Run]
    runs.forEach((r, ri) => {
      const o: PptxGenJS.TextPropsOptions = runProps(r, lang, canvas, p.kind === 'code')
      if (elementLink && !r.link) o.hyperlink = linkProps(elementLink)
      if (ri === 0) Object.assign(o, paragraphProps(p, canvas))
      if (ri > 0 && runs[ri - 1].breakAfter) o.softBreakBefore = true
      if (ri === runs.length - 1 && pi < paragraphs.length - 1) o.breakLine = true
      out.push({ text: sanitizeXmlText(r.text), options: o })
    })
  })
  return out
}

function paragraphProps(p: Paragraph, canvas: Canvas): PptxGenJS.TextPropsOptions {
  const o: PptxGenJS.TextPropsOptions = { align: p.align }
  if (p.kind === 'bullet') {
    o.bullet = { characterCode: '25AA' }
    o.indentLevel = p.level
  } else if (p.kind === 'number') {
    o.bullet = { type: 'number', ...(p.numberStart && p.numberStart !== 1 ? { numberStartAt: p.numberStart } : {}) }
    o.indentLevel = p.level
  }
  if (p.spaceBefore) o.paraSpaceBefore = pt(p.spaceBefore, canvas)
  if (p.spaceAfter) o.paraSpaceAfter = pt(p.spaceAfter, canvas)
  const size = p.runs[0]?.size
  if (size && p.lineHeight) o.lineSpacingMultiple = Math.max(1, Math.round((p.lineHeight / size) * 1000) / 1000)
  return o
}

function runProps(r: Run, lang: string, canvas: Canvas, code: boolean): PptxGenJS.TextPropsOptions {
  const o: PptxGenJS.TextPropsOptions = {
    fontFace: fontFor(code || r.code ? 'code' : 'body'),
    fontSize: pt(r.size, canvas),
    color: hex(r.color),
    lang,
  }
  if (r.bold) o.bold = true
  if (r.italic) o.italic = true
  if (r.underline) o.underline = { style: 'sng' }
  if (r.strike) o.strike = 'sngStrike'
  if (r.transparency) o.transparency = Math.round(r.transparency)
  if (r.highlight) o.highlight = hex(r.highlight)
  if (r.sup) o.superscript = true
  if (r.sub) o.subscript = true
  if (r.charSpacing) o.charSpacing = pt(r.charSpacing, canvas)
  if (r.link) o.hyperlink = linkProps(r.link)
  return o
}

function linkProps(link: NonNullable<Run['link']>): PptxGenJS.HyperlinkProps {
  return 'url' in link ? { url: sanitizeXmlText(link.url) } : { slide: link.slide }
}

function applyFrame(pptx: PptxGenJS, o: PptxGenJS.TextPropsOptions, frame: TextElement['frame'], canvas: Canvas): void {
  // §4.2: frame.transparency は fill.transparency にだけ反映する（run には収集時に掛け込んである）
  if (frame.fill) {
    const t = frame.fill.transparency ?? frame.transparency
    o.fill = { color: hex(frame.fill.color), ...(t ? { transparency: Math.round(t) } : {}) }
  }
  if (frame.line) o.line = { color: hex(frame.line.color), width: pt(frame.line.width, canvas), dashType: toDashType(frame.line.dash) }
  if (frame.radius) {
    o.shape = pptx.ShapeType.roundRect
    o.rectRadius = inch(frame.radius, canvas)
  }
}

// ---------------------------------------------------------------- 図形・線

function addShape(pptx: PptxGenJS, slide: PptxGenJS.Slide, e: ShapeElement, box: Box, name: string, lang: string, canvas: Canvas): void {
  const shapeType = (pptx.ShapeType as unknown as Record<string, PptxGenJS.SHAPE_NAME>)[e.shape] ?? pptx.ShapeType.rect
  const common: PptxGenJS.ShapeProps = { ...boxProps(box, canvas), objectName: name }
  if (e.frame.fill) common.fill = { color: hex(e.frame.fill.color), ...(e.frame.fill.transparency ? { transparency: Math.round(e.frame.fill.transparency) } : {}) }
  else common.fill = { color: 'FFFFFF', transparency: 100 }
  if (e.frame.line) common.line = { color: hex(e.frame.line.color), width: pt(e.frame.line.width, canvas), dashType: toDashType(e.frame.line.dash) }
  else common.line = { color: '000000', width: 0, transparency: 100 }
  if (e.frame.radius && e.shape === 'roundRect') common.rectRadius = inch(e.frame.radius, canvas)
  if (e.rotate) common.rotate = e.rotate

  const hasText = e.paragraphs?.some((p) => p.runs.some((r) => r.text.length))
  if (hasText) {
    // 要素リンクは addText に渡すと rels に登録されない（rIdundefined）ので、全 run に付ける（§4.3）
    const runs = toRuns(e.paragraphs!, lang, canvas, e.link)
    // 自由配置なので objectName は効く（placeholder 指定のときだけ捨てられる）
    const o: PptxGenJS.TextPropsOptions = { ...common, shape: shapeType, valign: e.valign ?? 'middle', margin: textMargin(e.frame.inset, canvas), lang, fit: 'shrink' }
    slide.addText(runs, o)
  } else {
    if (e.link) common.hyperlink = linkProps(e.link)
    slide.addShape(shapeType, common)
  }
}

function addLine(pptx: PptxGenJS, slide: PptxGenJS.Slide, e: LineElement, name: string, canvas: Canvas): void {
  // 端点は build の前処理でキャンバスの中に寄せてある（負の EMU を出さない）
  const x = Math.min(e.from.x, e.to.x)
  const y = Math.min(e.from.y, e.to.y)
  const w = Math.abs(e.to.x - e.from.x)
  const h = Math.abs(e.to.y - e.from.y)
  const flipV = (e.from.x < e.to.x && e.from.y > e.to.y) || (e.from.x > e.to.x && e.from.y < e.to.y)
  const line: PptxGenJS.ShapeLineProps = { color: hex(e.line.color), width: Math.max(0.5, pt(e.line.width, canvas)), dashType: toDashType(e.line.dash) }
  if (e.line.head) line.beginArrowType = arrow(e.line.head)
  if (e.line.tail) line.endArrowType = arrow(e.line.tail)
  slide.addShape(pptx.ShapeType.line, { x: emu(x, canvas), y: emu(y, canvas), w: emu(w, canvas), h: emu(h, canvas), line, flipV, objectName: name })
}

function arrow(v: string): PptxGenJS.ShapeLineProps['endArrowType'] {
  return (['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'].includes(v) ? v : 'arrow') as PptxGenJS.ShapeLineProps['endArrowType']
}

// ---------------------------------------------------------------- 画像

function addImage(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  e: ImageElement,
  box: Box,
  name: string,
  canvas: Canvas,
  sc: SlideCapture,
  opts: BuildOptions,
  report: Report,
): void {
  const key = e.captureId ?? e.src ?? ''
  let asset = assetOf(opts.assets, key)
  if (!asset && e.src?.startsWith('data:')) asset = { data: e.src }
  if (!asset) {
    report.warnings.push({ code: 'W-IMAGE', elementId: e.id, slide: sc.no, name, message: `画像を取得できない: ${e.captureId ? '撮影に失敗' : e.src}` })
    slide.addShape(pptx.ShapeType.rect, { ...boxProps(box, canvas), fill: { color: 'BFBFBF' }, line: { color: '999999', width: 0.5 }, objectName: name })
    return
  }
  const o: PptxGenJS.ImageProps = { data: asset.data, ...boxProps(box, canvas), objectName: name }
  if (e.alt) o.altText = sanitizeXmlText(e.alt)
  if (e.fit === 'contain' || e.fit === 'cover') o.sizing = { type: e.fit, w: emu(box.w, canvas), h: emu(box.h, canvas) }
  if (e.link) o.hyperlink = linkProps(e.link)
  slide.addImage(o)
  if (e.captureId) {
    report.replacements.push({
      slide: sc.no,
      elementId: e.id,
      name,
      reason: e.reason,
      selector: asset.selector ?? '',
      width: asset.width ?? Math.round(box.w * 2),
      height: asset.height ?? Math.round(box.h * 2),
    })
  }
}

// ---------------------------------------------------------------- 表

function addTable(slide: PptxGenJS.Slide, e: TableElement, box: Box, name: string, lang: string, canvas: Canvas): void {
  const rows: PptxGenJS.TableRow[] = e.rows.map((row) => row.map((cell) => toCell(cell, lang, canvas)))
  // 表の高さは rowH で決まる。寄せた枠（box.h）より高いなら、行の高さを比例で縮める
  const total = e.rowH.reduce((a, b) => a + b, 0)
  const scale = total > box.h && total > 0 ? box.h / total : 1
  slide.addTable(rows, {
    x: emu(box.x, canvas),
    y: emu(box.y, canvas),
    w: emu(box.w, canvas),
    h: emu(Math.min(box.h, total * scale), canvas),
    colW: e.colW.map((w) => emu(w, canvas)),
    rowH: e.rowH.map((h) => emu(h * scale, canvas)),
    objectName: name,
    fontFace: fontFor('body'),
    lang,
  } as PptxGenJS.TableProps)
}

function toCell(cell: Cell, lang: string, canvas: Canvas): PptxGenJS.TableCell {
  const runs = toRuns(cell.paragraphs, lang, canvas)
  const border = cell.border.map((b): PptxGenJS.BorderProps => (b.width > 0 ? { type: 'solid', pt: Math.max(0.25, pt(b.width, canvas)), color: hex(b.color) } : { type: 'none' })) as [
    PptxGenJS.BorderProps,
    PptxGenJS.BorderProps,
    PptxGenJS.BorderProps,
    PptxGenJS.BorderProps,
  ]
  const options: PptxGenJS.TableCellProps = {
    align: cell.align,
    valign: cell.valign,
    border,
    margin: cellMargin(cell.inset, canvas),
    lang,
  }
  if (cell.colspan && cell.colspan > 1) options.colspan = cell.colspan
  if (cell.rowspan && cell.rowspan > 1) options.rowspan = cell.rowspan
  if (cell.fill) options.fill = { color: hex(cell.fill) }
  return { text: runs, options }
}

// ---------------------------------------------------------------- 共通

function boxProps(box: Box, canvas: Canvas): { x: number; y: number; w: number; h: number } {
  return { x: emu(box.x, canvas), y: emu(box.y, canvas), w: emu(box.w, canvas), h: emu(box.h, canvas) }
}

function assetOf(assets: Record<string, string | Asset>, key: string): Asset | undefined {
  const v = assets[key]
  if (!v) return undefined
  return typeof v === 'string' ? { data: v } : v
}

/** '#rrggbb' / '#rgb' / 'rgb(a)(…)' → 'RRGGBB'（透明は 'FFFFFF'） */
export function hex(color: string): string {
  const c = (color ?? '').trim()
  let m = /^#([0-9a-f]{6})/i.exec(c)
  if (m) return m[1].toUpperCase()
  m = /^#([0-9a-f]{3})$/i.exec(c)
  if (m) return m[1].split('').map((x) => x + x).join('').toUpperCase()
  m = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\s*\)/i.exec(c)
  if (m) {
    if (m[4] !== undefined && Number(m[4]) === 0) return 'FFFFFF'
    return [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('').toUpperCase()
  }
  return 'FFFFFF'
}

function pptxVersion(): string {
  try {
    return (PptxGenJS as unknown as { version?: string }).version ?? (new PptxGenJS() as unknown as { version?: string }).version ?? ''
  } catch {
    return ''
  }
}
