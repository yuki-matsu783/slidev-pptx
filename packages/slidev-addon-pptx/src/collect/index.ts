// ブラウザで動く収集器（ppt-components.md §2、§3）。Vue にも Node にも依存しない。
// Playwright は関数を文字列化して page.evaluate に渡すので、collect はモジュール先頭の定数やヘルパを参照できない。
// すべて関数の中に閉じる。単位は Slidev キャンバス px、色は #rrggbb、透明度は 0–100。判断（レイアウト・対応表）はしない。
import type { Box, Capture, Cell, Dash, Element, FrameStyle, ImageElement, Paragraph, Run, SlideCapture, TextElement, Warning } from '../types.ts'

export function collect(): Capture {
  // ---------------------------------------------------------------- 道具
  const CONTAINER_TAGS = new Set(['DIV', 'SECTION', 'ARTICLE', 'MAIN', 'ASIDE', 'HEADER', 'FOOTER', 'SPAN', 'FIGURE', 'A', 'SMALL', 'STRONG', 'EM', 'B', 'I', 'U', 'S', 'DEL', 'MARK', 'SUB', 'SUP', 'CODE', 'KBD', 'ABBR', 'CITE', 'Q', 'LABEL', 'TIME'])
  const INLINE_TAGS = new Set(['A', 'SPAN', 'STRONG', 'EM', 'B', 'I', 'U', 'S', 'DEL', 'MARK', 'SUB', 'SUP', 'CODE', 'KBD', 'ABBR', 'CITE', 'Q', 'SMALL', 'BR', 'LABEL', 'TIME', 'WBR'])
  const FLOW_TAGS = new Set(['H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'P', 'UL', 'OL'])
  const REPLACE_TAGS = new Set(['SVG', 'CANVAS', 'IFRAME', 'VIDEO', 'AUDIO', 'OBJECT', 'EMBED'])
  const UI_SELECTOR = '.slidev-code-copy, .slidev-icon, .slidev-icon-btn, .slidev-nav, .slidev-page-root > .slidev-note'

  const round = (n: number) => Math.round(n * 100) / 100
  const px = (v: string) => parseFloat(v) || 0
  const hex = (color: string): string => {
    const m = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\s*\)/i.exec(color ?? '')
    if (!m) {
      if (/^#[0-9a-f]{6}$/i.test(color)) return color.toLowerCase()
      const s = /^#([0-9a-f]{3})$/i.exec(color ?? '')
      if (s) return '#' + s[1].toLowerCase().split('').map((c) => c + c).join('')
      return '#ffffff'
    }
    if (m[4] !== undefined && Number(m[4]) === 0) return '#ffffff'
    return '#' + [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('')
  }
  const alphaOf = (color: string): number => {
    const m = /^rgba?\(\s*\d+[,\s]+\d+[,\s]+\d+(?:[,\s/]+([\d.]+))?\s*\)/i.exec(color ?? '')
    if (!m) return color && color !== 'transparent' ? 1 : 0
    return m[1] === undefined ? 1 : Number(m[1])
  }
  const cs = (el: globalThis.Element) => getComputedStyle(el)

  // ---------------------------------------------------------------- 入口
  let containers = Array.from(document.querySelectorAll<HTMLElement>('#print-content > .print-slide-container'))
  if (!containers.length) containers = Array.from(document.querySelectorAll<HTMLElement>('.print-slide-container'))
  const first = containers[0]?.getBoundingClientRect()
  const canvas = { width: round(first?.width ?? 980), height: round(first?.height ?? 552) }
  const slides: SlideCapture[] = []

  for (const container of containers) {
    const page = container.querySelector<HTMLElement>('[data-slidev-no]')
    if (!page) continue
    const no = Number(page.dataset.slidevNo)
    const lang = page.getAttribute('lang') ?? ''
    const zoomRaw = cs(page).scale
    const zoom = zoomRaw && zoomRaw !== 'none' ? parseFloat(zoomRaw) || 1 : 1
    const crect = container.getBoundingClientRect()
    const elements: Element[] = []
    const warnings: Warning[] = []
    const dropped: Record<string, number> = {}
    let seq = 0
    const nextId = () => `s${no}-e${++seq}`
    const warn = (code: string, message: string, elementId?: string) => warnings.push({ code, message, elementId })
    const drop = (key: string) => {
      dropped[key] = (dropped[key] ?? 0) + 1
    }

    const boxOf = (el: globalThis.Element): Box => {
      const r = el.getBoundingClientRect()
      return { x: round(r.left - crect.left), y: round(r.top - crect.top), w: round(r.width), h: round(r.height) }
    }
    const len = (v: string) => round(px(v) * zoom) // computed の長さには zoom を掛ける（rect は掛けない）

    /** 祖先（スライドの page まで）の opacity の積 */
    const effectiveOpacity = (el: globalThis.Element | null): number => {
      let o = 1
      for (let e = el; e && e !== page.parentElement; e = e.parentElement) o *= parseFloat(cs(e).opacity) || 0
      return o
    }
    const isHidden = (el: globalThis.Element): 'none' | 'hidden' | 'opacity' | 'offcanvas' | false => {
      const s = cs(el)
      if (s.display === 'none') return 'none'
      if (s.visibility === 'hidden') return 'hidden'
      const b = boxOf(el)
      const tag = el.tagName
      // display: inline の要素がブロックの子（img など）だけを持つと自身の rect は 0 になる。子があれば子に判断を委ねる
      if (b.w === 0 && (b.h === 0 || tag !== 'HR') && el.children.length === 0) return 'none'
      if (effectiveOpacity(el) === 0) return 'opacity'
      if (b.x + b.w <= 0 || b.y + b.h <= 0 || b.x >= canvas.width || b.y >= canvas.height) return 'offcanvas'
      return false
    }

    interface Decor {
      fill?: FrameStyle['fill']
      line?: FrameStyle['line']
      radius?: number
      inset: [number, number, number, number]
      dropped: string[]
      decorated: boolean
      transform: boolean
    }
    const decorOf = (el: globalThis.Element): Decor => {
      const s = cs(el)
      const d: Decor = { inset: [len(s.paddingTop), len(s.paddingRight), len(s.paddingBottom), len(s.paddingLeft)], dropped: [], decorated: false, transform: false }
      const bgAlpha = alphaOf(s.backgroundColor)
      if (bgAlpha > 0) {
        d.fill = { color: hex(s.backgroundColor) }
        if (bgAlpha < 1) d.fill.transparency = Math.round((1 - bgAlpha) * 100)
        d.decorated = true
      }
      const widths = [s.borderTopWidth, s.borderRightWidth, s.borderBottomWidth, s.borderLeftWidth].map(px)
      if (widths.some((w) => w > 0)) {
        d.decorated = true
        const uniform = widths.every((w) => w === widths[0]) && [s.borderTopColor, s.borderRightColor, s.borderBottomColor, s.borderLeftColor].every((c) => c === s.borderTopColor)
        if (uniform && s.borderTopStyle !== 'none') {
          const dash: Dash = s.borderTopStyle === 'dashed' ? 'dash' : s.borderTopStyle === 'dotted' ? 'dot' : 'solid'
          d.line = { color: hex(s.borderTopColor), width: round(widths[0] * zoom), dash }
        } else d.dropped.push('border')
      }
      const radii = [s.borderTopLeftRadius, s.borderTopRightRadius, s.borderBottomRightRadius, s.borderBottomLeftRadius].map(px)
      if (radii.some((r) => r > 0)) {
        if (radii.every((r) => r === radii[0])) d.radius = round(radii[0] * zoom)
        else d.dropped.push('border-radius')
      }
      if (s.boxShadow && s.boxShadow !== 'none') {
        d.decorated = true
        d.dropped.push('box-shadow')
      }
      if (s.transform && s.transform !== 'none') {
        d.decorated = true
        d.transform = true
        d.dropped.push('transform')
      }
      if (s.filter && s.filter !== 'none') d.dropped.push('filter')
      if (s.backgroundImage && s.backgroundImage !== 'none' && !/url\(/.test(s.backgroundImage)) d.dropped.push('background-image')
      return d
    }
    const warnCss = (id: string, props: string[]) => {
      if (props.length) warn('W-CSS', `再現できない装飾を捨てた: ${props.join(', ')}`, id)
    }

    // ---------------------------------------------------------------- run
    const runStyle = (el: globalThis.Element, inCode: boolean, code: boolean): Omit<Run, 'text'> => {
      const s = cs(el)
      const r: Omit<Run, 'text'> = {
        size: len(s.fontSize),
        color: hex(s.color),
        bold: parseInt(s.fontWeight, 10) >= 600 || s.fontWeight === 'bold' || s.fontWeight === 'bolder',
        italic: s.fontStyle === 'italic' || s.fontStyle === 'oblique',
        underline: /underline/.test(s.textDecorationLine) && !el.closest('a'),
        strike: /line-through/.test(s.textDecorationLine) || !!el.closest('del, s'),
        code,
      }
      if (!inCode) {
        const o = effectiveOpacity(el)
        if (o < 1) r.transparency = Math.round((1 - o) * 100 * 1e6) / 1e6
      }
      if (code || el.closest('mark')) {
        const bgEl = (el.closest('code, kbd, mark') ?? el) as globalThis.Element
        const bg = cs(bgEl).backgroundColor
        if (alphaOf(bg) > 0) r.highlight = hex(bg)
      }
      if (el.closest('sup')) r.sup = true
      if (el.closest('sub')) r.sub = true
      const a = el.closest('a[href]') as HTMLAnchorElement | null
      if (a) r.link = linkOf(a)
      if (s.letterSpacing && s.letterSpacing !== 'normal') r.charSpacing = len(s.letterSpacing)
      if (s.textTransform && s.textTransform !== 'none') r.transparency = r.transparency // 値はそのまま。警告は呼び出し側
      return r
    }
    const linkOf = (a: HTMLAnchorElement): Run['link'] => {
      const href = a.getAttribute('href') ?? ''
      const m = /^#{1,2}(\d+)$/.exec(href) ?? /^\/(\d+)$/.exec(href)
      if (m) return { slide: Number(m[1]) }
      // a.href は正規化で末尾に / が付く。書いたままの URL を残し、相対だけ解く
      return { url: /^[a-z][a-z0-9+.-]*:/i.test(href) ? href : new URL(href, location.href).href }
    }

    /** inline の子を歩いて run を集める。ブロックの子に当たったら止めて残りを返す */
    const collectRuns = (parent: globalThis.Element, id: string, inCode: boolean, out: Run[] = []): Run[] => {
      for (const node of Array.from(parent.childNodes)) {
        if (node.nodeType === 3) {
          const raw = node.nodeValue ?? ''
          const text = inCode ? raw : raw.replace(/\s+/g, ' ')
          const holder = (node.parentElement ?? parent) as globalThis.Element
          if (!text.trim() && !inCode) {
            // 空白だけのノードは、前の run の文字を変えず、親の書式で独立した run にする（枠の端では trimRuns が落とす）
            if (out.length && text.length && !out[out.length - 1].text.endsWith(' ')) out.push({ text: ' ', ...runStyle(parent, inCode, false) })
            continue
          }
          const code = !!holder.closest('code, kbd')
          out.push({ text, ...runStyle(holder, inCode, code) })
          continue
        }
        if (node.nodeType !== 1) continue
        const el = node as HTMLElement
        if (el.tagName === 'BR') {
          if (out.length) out[out.length - 1].breakAfter = true
          continue
        }
        if (el.matches(UI_SELECTOR)) continue
        if (el.classList.contains('katex')) {
          const html = el.querySelector('.katex-html')
          const text = (html?.textContent ?? el.textContent ?? '').replace(/\s+/g, ' ').trim()
          out.push({ text, ...runStyle(el, inCode, false), italic: true })
          warn('W-MATH-INLINE', `インライン数式 "${text}" を文字に平坦化した`, id)
          continue
        }
        if (cs(el).display === 'none') continue
        if (INLINE_TAGS.has(el.tagName) || cs(el).display.startsWith('inline')) {
          const tt = cs(el).textTransform
          if (tt && tt !== 'none') warn('W-CSS', `再現できない装飾を捨てた: text-transform`, id)
          collectRuns(el, id, inCode, out)
          continue
        }
        // ブロックの子（div など）: 中の inline を続けて拾う
        collectRuns(el, id, inCode, out)
      }
      return out
    }
    const trimRuns = (runs: Run[]): Run[] => {
      while (runs.length && !runs[0].text.trim() && !runs[0].breakAfter) runs.shift()
      while (runs.length && !runs[runs.length - 1].text.trim() && !runs[runs.length - 1].breakAfter) runs.pop()
      if (runs.length) {
        runs[0].text = runs[0].text.replace(/^\s+/, '')
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '')
      }
      return runs.filter((r) => r.text.length || r.breakAfter)
    }

    const alignOf = (el: globalThis.Element): Paragraph['align'] => {
      const a = cs(el).textAlign
      if (a === 'center') return 'center'
      if (a === 'right' || a === 'end') return 'right'
      if (a === 'justify') return 'justify'
      return 'left'
    }
    const paragraphOf = (el: globalThis.Element, kind: Paragraph['kind'], level: number, runs: Run[]): Paragraph => {
      const s = cs(el)
      const size = runs[0]?.size || len(s.fontSize)
      const lh = s.lineHeight === 'normal' ? round(size * 1.2) : len(s.lineHeight)
      return { kind, level, align: alignOf(el), lineHeight: lh, spaceBefore: len(s.marginTop), spaceAfter: len(s.marginBottom), runs }
    }

    // ---------------------------------------------------------------- 段落（ブロック → Paragraph[]）
    const blockParagraphs = (el: globalThis.Element, id: string): Paragraph[] => {
      const tag = el.tagName
      if (/^H[1-6]$/.test(tag)) return [paragraphOf(el, 'heading', Number(tag[1]), trimRuns(collectRuns(el, id, false)))]
      if (tag === 'UL' || tag === 'OL') return listParagraphs(el, id, 0)
      if (tag === 'PRE') return codeParagraphs(el)
      return [paragraphOf(el, 'plain', 0, trimRuns(collectRuns(el, id, false)))]
    }
    const listParagraphs = (list: globalThis.Element, id: string, level: number): Paragraph[] => {
      const out: Paragraph[] = []
      const kind: Paragraph['kind'] = list.tagName === 'OL' ? 'number' : 'bullet'
      const start = list.tagName === 'OL' ? Number(list.getAttribute('start') ?? '1') : undefined
      let n = 0
      for (const li of Array.from(list.children).filter((c) => c.tagName === 'LI')) {
        n++
        let firstDone = false
        const pushPara = (src: globalThis.Element, runs: Run[]) => {
          const p = paragraphOf(src, firstDone ? 'plain' : kind, level, runs)
          if (!firstDone && start !== undefined) p.numberStart = start + n - 1
          firstDone = true
          out.push(p)
        }
        // li 直下の inline とテキストをまとめる。ブロックの子（p, ul, ol, pre, table…）で区切る
        let pending: Run[] = []
        const flush = () => {
          const runs = trimRuns(pending)
          pending = []
          if (runs.length) pushPara(li, runs)
        }
        for (const node of Array.from(li.childNodes)) {
          if (node.nodeType === 3 || (node.nodeType === 1 && (INLINE_TAGS.has((node as globalThis.Element).tagName) || cs(node as globalThis.Element).display.startsWith('inline')))) {
            const tmp = document.createDocumentFragment()
            void tmp
            const holder = document.createElement('span')
            // 直接 collectRuns を node 単位で
            if (node.nodeType === 3) {
              const text = (node.nodeValue ?? '').replace(/\s+/g, ' ')
              if (text.trim()) pending.push({ text, ...runStyle(li, false, false) })
              else if (pending.length && text.length && !pending[pending.length - 1].text.endsWith(' ')) pending.push({ text: ' ', ...runStyle(li, false, false) })
            } else {
              const el = node as globalThis.Element
              if (el.tagName === 'BR') {
                if (pending.length) pending[pending.length - 1].breakAfter = true
              } else if (el.classList.contains('katex')) {
                collectRuns(holder, id, false, pending) // no-op holder
                const html = el.querySelector('.katex-html')
                const text = (html?.textContent ?? '').replace(/\s+/g, ' ').trim()
                pending.push({ text, ...runStyle(el, false, false), italic: true })
                warn('W-MATH-INLINE', `インライン数式 "${text}" を文字に平坦化した`, id)
              } else collectRuns(el, id, false, pending)
            }
            continue
          }
          if (node.nodeType !== 1) continue
          const el = node as globalThis.Element
          flush()
          if (el.tagName === 'P' || /^H[1-6]$/.test(el.tagName)) pushPara(el, trimRuns(collectRuns(el, id, false)))
          else if (el.tagName === 'UL' || el.tagName === 'OL') {
            firstDone = true
            out.push(...listParagraphs(el, id, level + 1))
          } else if (el.tagName === 'PRE') {
            firstDone = true
            out.push(...codeParagraphs(el))
          } else if (el.tagName === 'DIV' || el.tagName === 'SECTION') {
            pushPara(el, trimRuns(collectRuns(el, id, false)))
          } else {
            warn('W-LI-BLOCK', `箇条書きの中の ${el.tagName.toLowerCase()} は出せないので無視した`, id)
          }
        }
        flush()
        if (!firstDone) pushPara(li, [])
      }
      return out
    }
    const codeParagraphs = (pre: globalThis.Element): Paragraph[] => {
      const s = cs(pre)
      const size = len(s.fontSize)
      const color = hex(s.color)
      const lh = s.lineHeight === 'normal' ? round(size * 1.2) : len(s.lineHeight)
      const lines = Array.from(pre.querySelectorAll('.line'))
      const texts = lines.length ? lines.map((l) => l.textContent ?? '') : (pre.textContent ?? '').replace(/\n$/, '').split('\n')
      if (pre.querySelector('.line span[style], .line [class*="shiki"]') || (pre.className && /shiki/.test(pre.className))) drop('code-highlight')
      return texts.map((t) => ({ kind: 'code' as const, level: 0, align: 'left' as const, lineHeight: lh, spaceBefore: 0, spaceAfter: 0, runs: [{ text: t, size, color, bold: false, italic: false, underline: false, strike: false, code: true }] }))
    }

    // ---------------------------------------------------------------- 枠（流し込み）
    interface FlowState {
      region: 'root' | 'left' | 'right'
      blocks: { el: globalThis.Element; paragraphs: Paragraph[] }[]
      /** この枠が title 直後（区切り無し）に開いたか */
      afterTitle: boolean
      /** 右列で最初の枠か */
      firstInRight: boolean
    }
    let titleFound = false
    let titleJustSeen = false // 直近に title 候補の枠を閉じ、まだ区切りブロックを見ていない
    let state: FlowState | null = null
    let regionSawSeparator = false

    const openFlow = (region: FlowState['region']): FlowState => {
      if (!state) state = { region, blocks: [], afterTitle: titleJustSeen && region !== 'right', firstInRight: region === 'right' && !regionSawSeparator }
      return state
    }
    const closeFlow = () => {
      if (!state || !state.blocks.length) {
        state = null
        return
      }
      const id = nextId()
      const rects = state.blocks.map((b) => boxOf(b.el))
      const x = Math.min(...rects.map((r) => r.x))
      const y = Math.min(...rects.map((r) => r.y))
      const x2 = Math.max(...rects.map((r) => r.x + r.w))
      const y2 = Math.max(...rects.map((r) => r.y + r.h))
      const box = { x, y, w: round(x2 - x), h: round(y2 - y) }
      const t: TextElement = {
        id,
        name: '',
        source: 'markdown',
        kind: 'text',
        box,
        boxSource: 'measured',
        paragraphs: state.blocks.flatMap((b) => b.paragraphs),
        frame: { inset: [0, 0, 0, 0] },
        fit: { contentHeight: box.h, boxHeight: box.h },
        valign: 'top',
      }
      if (state.afterTitle) t.roleHint = 'body'
      else if (state.firstInRight) t.roleHint = 'body2'
      elements.push(t)
      // 段落ごとの W-* は id を持てないので枠の id に付け直す
      for (const w of warnings) if (w.elementId === '?') w.elementId = id
      titleJustSeen = false
      state = null
    }
    const separator = () => {
      closeFlow()
      titleJustSeen = false
      regionSawSeparator = true
    }

    const pushHeadingAsTitle = (el: globalThis.Element) => {
      closeFlow()
      const id = nextId()
      const t: TextElement = {
        id,
        name: '',
        source: 'markdown',
        kind: 'text',
        roleHint: 'title',
        box: boxOf(el),
        boxSource: 'measured',
        paragraphs: blockParagraphs(el, id),
        frame: { inset: [0, 0, 0, 0] },
        fit: { contentHeight: round(el.scrollHeight * 1), boxHeight: round(el.clientHeight * 1) },
        valign: 'top',
      }
      elements.push(t)
      titleFound = true
      titleJustSeen = true
    }

    const pushFrame = (el: globalThis.Element, paragraphs: Paragraph[], frame: FrameStyle, id: string, source: TextElement['source'] = 'markdown', extra: Partial<TextElement> = {}) => {
      const t: TextElement = {
        id,
        name: '',
        source,
        kind: 'text',
        box: boxOf(el),
        boxSource: 'measured',
        paragraphs,
        frame,
        fit: { contentHeight: round(el.scrollHeight * zoom), boxHeight: round(el.clientHeight * zoom) },
        valign: 'top',
        ...extra,
      }
      elements.push(t)
      return t
    }
    const pushImage = (el: globalThis.Element, partial: Partial<ImageElement>, id = nextId()): ImageElement => {
      const img: ImageElement = { id, name: '', source: 'markdown', kind: 'image', box: boxOf(el), boxSource: 'measured', ...partial }
      el.setAttribute('data-ppt-capture-id', id)
      elements.push(img)
      return img
    }
    const replace = (el: globalThis.Element, reason: ImageElement['reason'], id = nextId()): ImageElement => pushImage(el, { source: 'replaced', captureId: id, reason }, id)

    // ---------------------------------------------------------------- 表
    const tableElement = (tbl: globalThis.Element, id: string) => {
      const trs = Array.from(tbl.querySelectorAll('tr')).filter((tr) => tr.closest('table') === tbl)
      const headerRows = Array.from(tbl.querySelectorAll('thead > tr')).filter((tr) => tr.closest('table') === tbl).length
      const tblS = cs(tbl)
      const rows: Cell[][] = trs.map((tr) => {
        const trS = cs(tr)
        return Array.from(tr.children)
          .filter((c) => c.tagName === 'TD' || c.tagName === 'TH')
          .map((td) => {
            const s = cs(td)
            const side = (w: string, c: string, fw: string, fc: string) => {
              const width = px(w) > 0 ? round(px(w) * zoom) : px(fw) > 0 ? round(px(fw) * zoom) : 0
              return { width, color: width ? hex(px(w) > 0 ? c : fc) : '' }
            }
            const cell: Cell = {
              paragraphs: [paragraphOf(td, 'plain', 0, trimRuns(collectRuns(td, id, false)))],
              border: [
                side(s.borderTopWidth, s.borderTopColor, trS.borderTopWidth, trS.borderTopColor),
                side(s.borderRightWidth, s.borderRightColor, tblS.borderRightWidth, tblS.borderRightColor),
                side(s.borderBottomWidth, s.borderBottomColor, trS.borderBottomWidth, trS.borderBottomColor),
                side(s.borderLeftWidth, s.borderLeftColor, tblS.borderLeftWidth, tblS.borderLeftColor),
              ],
              inset: [len(s.paddingTop), len(s.paddingRight), len(s.paddingBottom), len(s.paddingLeft)],
              align: alignOf(td),
              valign: s.verticalAlign === 'top' ? 'top' : s.verticalAlign === 'bottom' ? 'bottom' : 'middle',
            }
            const cspan = Number(td.getAttribute('colspan') ?? '1')
            const rspan = Number(td.getAttribute('rowspan') ?? '1')
            if (cspan > 1) cell.colspan = cspan
            if (rspan > 1) cell.rowspan = rspan
            if (alphaOf(s.backgroundColor) > 0) cell.fill = hex(s.backgroundColor)
            return cell
          })
      })
      const firstRow = trs[0] ? Array.from(trs[0].children).filter((c) => c.tagName === 'TD' || c.tagName === 'TH') : []
      const colW = firstRow.map((c) => round(c.getBoundingClientRect().width))
      const rowH = trs.map((tr) => round(tr.getBoundingClientRect().height))
      elements.push({ id, name: '', source: 'markdown', kind: 'table', box: boxOf(tbl), boxSource: 'measured', colW, rowH, headerRows, rows })
    }

    // ---------------------------------------------------------------- PPT 部品
    const pptPart = (el: HTMLElement, region: FlowState['region']) => {
      separator()
      const id = nextId()
      const type = el.dataset.ppt
      const name = el.dataset.pptName ?? ''
      let opts: Record<string, unknown> = {}
      try {
        opts = el.dataset.pptOpts ? JSON.parse(el.dataset.pptOpts) : {}
      } catch {
        /* 無視 */
      }
      let propBox: Partial<Box> = {}
      try {
        propBox = el.dataset.pptBox ? JSON.parse(el.dataset.pptBox) : {}
      } catch {
        /* 無視 */
      }
      const measured = boxOf(el)
      const box: Box = { ...measured, ...propBox }
      const boxSource: Element['boxSource'] = Object.keys(propBox).length ? 'prop' : 'measured'
      const nested = Array.from(el.querySelectorAll('[data-ppt]'))
      for (const n of nested) warn('W-NESTED-PPT', `PPT 部品の中の PPT 部品（${(n as HTMLElement).dataset.ppt}）は無視した`, id)
      if (el.dataset.pptExport === 'image') {
        const img = replace(el, 'explicit', id)
        img.name = name
        img.box = box
        img.boxSource = boxSource
        return
      }
      const inner = (): Paragraph[] => {
        const out: Paragraph[] = []
        const walkInner = (parent: globalThis.Element) => {
          for (const c of Array.from(parent.children)) {
            if ((c as HTMLElement).dataset?.ppt !== undefined) continue
            if (FLOW_TAGS.has(c.tagName) || c.tagName === 'PRE') out.push(...blockParagraphs(c, id))
            else if (c.tagName === 'TABLE' || c.tagName === 'IMG' || c.tagName === 'BLOCKQUOTE' || c.classList.contains('katex-display')) warn('W-PPT-CONTENT', `PptText の中の ${c.tagName.toLowerCase()} は無視した`, id)
            else if (c.tagName === 'BR') {
              /* 段落の中で扱う */
            } else walkInner(c)
          }
          // 直下のテキスト
          const direct = Array.from(parent.childNodes).filter((n) => n.nodeType === 3 && (n.nodeValue ?? '').trim())
          if (direct.length) {
            const runs = trimRuns(collectRuns(parent, id, false).filter((r) => r.text.trim()))
            if (runs.length && !parent.querySelector('p, ul, ol, h1, h2, h3, h4, h5, h6')) out.push(paragraphOf(parent, 'plain', 0, runs))
          }
        }
        walkInner(el)
        return out
      }
      const d = decorOf(el)
      const frame: FrameStyle = { inset: d.inset }
      // 塗りと線は props（opts）を正とし、無ければ computed から（部品は props を CSS にも写しているので同じ値になる）
      if (typeof opts.fill === 'string' && opts.fill !== 'none') frame.fill = { color: hex(opts.fill as string) }
      else if (d.fill && opts.fill !== 'none') frame.fill = d.fill
      if (opts.line && opts.line !== 'none') {
        const l = opts.line as { color?: string; width?: number; dash?: Dash } | string
        frame.line = typeof l === 'string' ? { color: hex(l), width: 1, dash: 'solid' } : { color: hex(l.color ?? '#000000'), width: l.width ?? 1, dash: l.dash ?? 'solid' }
      } else if (d.line && opts.line !== 'none') frame.line = d.line
      if (typeof opts.radius === 'number') frame.radius = opts.radius
      else if (d.radius) frame.radius = d.radius
      if (typeof opts.padding === 'number') frame.inset = [opts.padding, opts.padding, opts.padding, opts.padding]
      else if (Array.isArray(opts.padding)) frame.inset = opts.padding as FrameStyle['inset']

      if (type === 'text') {
        const t = pushFrame(el, inner(), frame, id, 'ppt', { name, box, boxSource, valign: (opts.valign as TextElement['valign']) ?? 'top' })
        void t
        void region
        return
      }
      if (type === 'shape') {
        const shape = (opts.type as string) ?? 'rect'
        const lineOpts = opts.line as { color?: string; width?: number; dash?: Dash; head?: string; tail?: string } | string | undefined
        if (shape === 'line') {
          const from = { x: box.x, y: box.y }
          const to = { x: box.x + box.w, y: box.y + box.h }
          const l = typeof lineOpts === 'object' && lineOpts ? lineOpts : { color: typeof lineOpts === 'string' ? lineOpts : '#000000' }
          elements.push({ id, name, source: 'ppt', kind: 'line', box, boxSource, from, to, line: { color: hex(l.color ?? '#000000'), width: l.width ?? 1, dash: l.dash ?? 'solid', head: l.head, tail: l.tail } })
          return
        }
        if (!frame.fill && opts.fill !== 'none') frame.fill = { color: '#ffffff' }
        const paragraphs = inner()
        elements.push({
          id, name, source: 'ppt', kind: 'shape', shape, box, boxSource, frame,
          rotate: typeof opts.rotate === 'number' ? opts.rotate : undefined,
          paragraphs: paragraphs.length ? paragraphs : undefined,
          valign: (opts.valign as TextElement['valign']) ?? 'middle',
        })
        return
      }
      if (type === 'image') {
        const img = (el.tagName === 'IMG' ? el : el.querySelector('img')) as HTMLImageElement | null
        pushImage(el, { name, source: 'ppt', box, boxSource, src: img?.currentSrc || img?.src, alt: img?.alt || (opts.alt as string) || '', fit: (opts.fit as ImageElement['fit']) ?? 'contain' }, id)
        return
      }
      if (type === 'table') {
        const tbl = el.tagName === 'TABLE' ? el : el.querySelector('table')
        if (tbl) {
          tableElement(tbl, id)
          const last = elements[elements.length - 1]
          last.name = name
          last.box = box
          last.boxSource = boxSource
        }
        return
      }
      warn('W-NESTED-PPT', `不明な data-ppt="${type}"`, id)
    }

    // ---------------------------------------------------------------- 歩く（§2.2）
    const walk = (el: HTMLElement, region: FlowState['region'], throughTransparent: boolean) => {
      const tag = el.tagName
      // 0
      if (el.matches(UI_SELECTOR)) return
      // 1
      const hidden = isHidden(el)
      if (hidden) {
        if (hidden === 'opacity') warn('W-HIDDEN', 'opacity 0 の要素を飛ばした', undefined)
        else if (hidden === 'offcanvas') warn('W-HIDDEN', 'キャンバスの外の要素を飛ばした', undefined)
        return
      }
      // 2
      if (el.dataset.ppt !== undefined) {
        pptPart(el, region)
        return
      }
      // 3
      if (el.dataset.pptExport === 'image') {
        separator()
        replace(el, 'explicit')
        return
      }
      // 4
      const s = cs(el)
      if (tag !== 'IMG' && s.backgroundImage && s.backgroundImage !== 'none') {
        separator()
        const hasText = !!(el.textContent ?? '').trim()
        const m = /url\(["']?([^"')]+)["']?\)/.exec(s.backgroundImage)
        if (hasText || !m) {
          replace(el, hasText ? 'explicit' : 'gradient')
        } else {
          if (/gradient\(/.test(s.backgroundImage)) warn('W-CSS', '再現できない装飾を捨てた: background-image の gradient', undefined)
          pushImage(el, { src: new URL(m[1], location.href).href })
        }
        return
      }
      // 5
      if (el.classList.contains('katex-display') || el.classList.contains('mermaid') || REPLACE_TAGS.has(tag)) {
        separator()
        replace(el, el.classList.contains('katex-display') ? 'math' : el.classList.contains('mermaid') ? 'mermaid' : tag === 'SVG' ? 'svg' : 'unknown-element')
        return
      }
      if (tag === 'P') {
        const kids = Array.from(el.children)
        if (kids.length === 1 && kids[0].classList.contains('katex-display') && !(el.textContent ?? '').replace(kids[0].textContent ?? '', '').trim()) {
          separator()
          replace(el, 'math')
          return
        }
      }
      // 6
      if (tag === 'IMG') {
        separator()
        const img = el as HTMLImageElement
        const a = el.closest('a[href]') as HTMLAnchorElement | null
        const e = pushImage(el, { src: img.currentSrc || img.src, alt: img.alt || '' })
        if (a && a.closest('.slidev-layout, [data-slidev-no]')) e.link = linkOf(a)
        return
      }
      // 7
      if (tag === 'TABLE') {
        separator()
        tableElement(el, nextId())
        return
      }
      // 8
      if (tag === 'PRE') {
        separator()
        const id = nextId()
        const d = decorOf(el)
        const frame: FrameStyle = { inset: d.inset, fill: d.fill, radius: d.radius }
        pushFrame(el, codeParagraphs(el), frame, id)
        return
      }
      // 9
      if (tag === 'BLOCKQUOTE') {
        separator()
        const id = nextId()
        const d = decorOf(el)
        const paragraphs: Paragraph[] = []
        const ps = Array.from(el.children).filter((c) => FLOW_TAGS.has(c.tagName))
        if (ps.length) for (const p of ps) paragraphs.push(...blockParagraphs(p, id))
        else paragraphs.push(paragraphOf(el, 'plain', 0, trimRuns(collectRuns(el, id, false))))
        pushFrame(el, paragraphs, { inset: d.inset, fill: d.fill, radius: d.radius }, id)
        drop('blockquote-border')
        return
      }
      // 10
      if (tag === 'HR') {
        separator()
        const b = boxOf(el)
        const width = Math.max(1, round((px(s.borderTopWidth) || 1) * zoom))
        elements.push({ id: nextId(), name: '', source: 'markdown', kind: 'line', box: b, boxSource: 'measured', line: { color: hex(s.borderTopColor), width, dash: 'solid' }, from: { x: b.x, y: b.y }, to: { x: b.x + b.w, y: b.y } })
        return
      }
      // 11
      if (FLOW_TAGS.has(tag)) {
        if (/^H[12]$/.test(tag) && !titleFound && region !== 'right' && throughTransparent) {
          pushHeadingAsTitle(el)
          return
        }
        const st = openFlow(region)
        const id = '?'
        st.blocks.push({ el, paragraphs: blockParagraphs(el, id) })
        return
      }
      // 12〜15
      const container = CONTAINER_TAGS.has(tag)
      const d = decorOf(el)
      if (container && !d.decorated) {
        walkContainer(el, region, throughTransparent)
        return
      }
      if (container) {
        const separators = Array.from(el.querySelectorAll('*')).some((c) => isSeparatorLike(c as HTMLElement))
        if (!separators) {
          // 13: 自前のテキスト枠
          separator()
          const id = nextId()
          const paragraphs = containerParagraphs(el, id)
          const frame: FrameStyle = { inset: d.inset, fill: d.fill, line: d.line, radius: d.radius }
          const t = pushFrame(el, paragraphs, frame, id)
          if (d.transform) {
            const m = /rotate\(([-\d.]+)deg\)/.exec((el as HTMLElement).style.transform ?? '')
            if (m) t.rotate = Number(m[1])
          }
          warnCss(id, d.dropped)
          return
        }
        // 14: 装飾を捨てて中を歩く
        separator()
        warnCss(undefined as unknown as string, [...new Set([...d.dropped, ...(d.fill ? ['background'] : []), ...(d.line ? ['border'] : [])])])
        walkContainer(el, region, false)
        return
      }
      // 15
      separator()
      replace(el, 'unknown-element')
    }
    const isSeparatorLike = (c: HTMLElement): boolean => {
      const t = c.tagName
      if (['TABLE', 'PRE', 'BLOCKQUOTE', 'IMG', 'HR', 'BUTTON', 'INPUT', 'SELECT', 'TEXTAREA', 'DETAILS'].includes(t)) return true
      if (REPLACE_TAGS.has(t) || c.classList.contains('katex-display') || c.classList.contains('mermaid')) return true
      if (c.dataset.ppt !== undefined || c.dataset.pptExport === 'image') return true
      if (CONTAINER_TAGS.has(t) && decorOf(c).decorated) return true
      const bg = cs(c).backgroundImage
      if (bg && bg !== 'none') return true
      return false
    }
    /** 12: 透明な入れ物。直下のテキスト・inline は連なりごとに 1 段落として流し込みに数える */
    const walkContainer = (el: HTMLElement, region: FlowState['region'], throughTransparent: boolean) => {
      let pending: Run[] = []
      const flushInline = () => {
        const runs = trimRuns(pending)
        pending = []
        if (!runs.length) return
        const st = openFlow(region)
        st.blocks.push({ el, paragraphs: [paragraphOf(el, 'plain', 0, runs)] })
      }
      for (const node of Array.from(el.childNodes)) {
        if (node.nodeType === 3) {
          const text = (node.nodeValue ?? '').replace(/\s+/g, ' ')
          if (text.trim()) pending.push({ text, ...runStyle(el, false, false) })
          else if (pending.length && text.length && !pending[pending.length - 1].text.endsWith(' ')) pending.push({ text: ' ', ...runStyle(el, false, false) })
          continue
        }
        if (node.nodeType !== 1) continue
        const c = node as HTMLElement
        if (c.matches(UI_SELECTOR)) continue
        if (c.tagName === 'BR') {
          if (pending.length) pending[pending.length - 1].breakAfter = true
          continue
        }
        const inline = INLINE_TAGS.has(c.tagName) && !c.querySelector('img, table, pre, div, p, ul, ol') && c.dataset.ppt === undefined && c.dataset.pptExport !== 'image'
        if (inline && cs(c).display !== 'none') {
          if (c.classList.contains('katex')) {
            const html = c.querySelector('.katex-html')
            const text = (html?.textContent ?? '').replace(/\s+/g, ' ').trim()
            pending.push({ text, ...runStyle(c, false, false), italic: true })
            warn('W-MATH-INLINE', `インライン数式 "${text}" を文字に平坦化した`, '?')
          } else collectRuns(c, '?', false, pending)
          continue
        }
        flushInline()
        walk(c, region, throughTransparent)
      }
      flushInline()
    }
    /** 13 の中身: 装飾つきの箱の中を、透明な入れ物だけを通って段落にする */
    const containerParagraphs = (el: globalThis.Element, id: string): Paragraph[] => {
      const out: Paragraph[] = []
      let pending: Run[] = []
      const flush = () => {
        const runs = trimRuns(pending)
        pending = []
        if (runs.length) out.push(paragraphOf(el, 'plain', 0, runs))
      }
      const rec = (parent: globalThis.Element) => {
        for (const node of Array.from(parent.childNodes)) {
          if (node.nodeType === 3) {
            const text = (node.nodeValue ?? '').replace(/\s+/g, ' ')
            if (text.trim()) pending.push({ text, ...runStyle(parent, false, false) })
            else if (pending.length && text.length && !pending[pending.length - 1].text.endsWith(' ')) pending.push({ text: ' ', ...runStyle(parent, false, false) })
            continue
          }
          if (node.nodeType !== 1) continue
          const c = node as HTMLElement
          if (c.matches(UI_SELECTOR) || cs(c).display === 'none') continue
          if (c.tagName === 'BR') {
            if (pending.length) pending[pending.length - 1].breakAfter = true
            continue
          }
          if (INLINE_TAGS.has(c.tagName)) {
            collectRuns(c, id, false, pending)
            continue
          }
          flush()
          if (FLOW_TAGS.has(c.tagName) || c.tagName === 'PRE') out.push(...blockParagraphs(c, id))
          else {
            // 透明な入れ物: 自分の段落として
            const runs = trimRuns(collectRuns(c, id, false))
            if (c.querySelector('p, ul, ol, h1, h2, h3, h4, h5, h6, div')) rec(c)
            else if (runs.length) out.push(paragraphOf(c, 'plain', 0, runs))
          }
        }
      }
      rec(el)
      flush()
      return out
    }

    // ---------------------------------------------------------------- 根の列（§2.1）
    const roots: { el: HTMLElement; layout: boolean }[] = []
    const findRoots = (el: HTMLElement) => {
      for (const c of Array.from(el.children) as HTMLElement[]) {
        if (c.matches(UI_SELECTOR)) continue
        if (c.classList.contains('slidev-layout')) roots.push({ el: c, layout: true })
        else if (c.querySelector('.slidev-layout')) findRoots(c)
        else roots.push({ el: c, layout: false })
      }
    }
    findRoots(page)

    for (const root of roots) {
      regionSawSeparator = false
      state = null
      if (root.layout) {
        const left = root.el.querySelector<HTMLElement>(':scope > .col-left')
        const right = root.el.querySelector<HTMLElement>(':scope > .col-right')
        if (left && right) {
          walkContainer(left, 'left', true)
          closeFlow()
          regionSawSeparator = false
          titleJustSeen = false
          walkContainer(right, 'right', true)
          closeFlow()
        } else {
          walkContainer(root.el, 'root', true)
          closeFlow()
        }
      } else {
        walk(root.el, 'root', true)
        closeFlow()
      }
    }

    // W-HIDDEN などで elementId が '?' のまま残ったものは外す
    for (const w of warnings) if (w.elementId === '?') delete w.elementId

    const slide: SlideCapture & { dropped?: Record<string, number> } = {
      no,
      lang,
      zoom,
      backgroundColor: hex(cs(container).backgroundColor),
      elements,
      warnings,
    }
    if (Object.keys(dropped).length) slide.dropped = dropped
    slides.push(slide)
  }

  return { canvas, slides }
}
