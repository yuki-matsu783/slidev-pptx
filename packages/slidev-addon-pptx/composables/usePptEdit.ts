// 試作: PPT 部品の中身（slot の Markdown）を小窓で編集し、slides.md のその部品の中だけを書き換える。
// 部品は drag="id" で探す。書き戻しは Slidev の useDynamicSlideInfo().update（発表者ノートの編集と同じ口）。
import { useDynamicSlideInfo } from '@slidev/client/composables/useSlideInfo.ts'
import { injectionCurrentPage } from '@slidev/client/constants.ts'
import { computed, inject, ref, unref } from 'vue'

export type PptTag = 'PptText' | 'PptShape'

export interface Located {
  start: number
  end: number
  openTag: string
  inner: string
  selfClosing: boolean
}

// 属性の値の中の > で開始タグが切れないよう、引用符の中は丸ごと読み飛ばす
const ATTRS = `(?:[^>"']|"[^"]*"|'[^']*')*?`
const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/** スライドの本文から drag="id" を持つ <tag> を 1 つだけ探す。見つからない・複数あるときは理由の文字列 */
export function locate(content: string, tag: PptTag, id: string): Located | string {
  const hasId = new RegExp(`\\sdrag=(["'])${escapeRe(id)}\\1`)
  const hits: Located[] = []
  for (const m of content.matchAll(new RegExp(`<${tag}\\b${ATTRS}(\\/?)>`, 'g'))) {
    const openTag = m[0]
    if (!hasId.test(openTag)) continue
    const start = m.index!
    const openEnd = start + openTag.length
    if (m[1] === '/') {
      hits.push({ start, end: openEnd, openTag, inner: '', selfClosing: true })
      continue
    }
    const close = content.indexOf(`</${tag}>`, openEnd)
    if (close < 0) return `<${tag} drag="${id}"> の閉じタグが見つかりません`
    hits.push({ start, end: close + tag.length + 3, openTag, inner: content.slice(openEnd, close), selfClosing: false })
  }
  if (hits.length === 0) return `<${tag} drag="${id}"> がこのスライドの slides.md に見つかりません`
  if (hits.length > 1) return `<${tag} drag="${id}"> が ${hits.length} 個あります。drag の名前を分けてください`
  return hits[0]
}

/** 見つけた部品の中身を text に差し替えた本文を返す */
export function replaceInner(content: string, loc: Located, tag: PptTag, text: string): string {
  const body = text.trim()
  const open = loc.selfClosing ? loc.openTag.replace(/\s*\/>$/, '>') : loc.openTag
  // 図形の 1 行ラベル（Markdown の記号なし）はタグと同じ行に置く。それ以外は前後に空行を置いて Markdown として描かせる
  const inline = tag === 'PptShape' && !body.includes('\n') && !/[*_`[<]/.test(body)
  const inner = inline ? body : `\n\n${body}\n\n`
  return content.slice(0, loc.start) + open + inner + `</${tag}>` + content.slice(loc.end)
}

export function usePptEdit(props: { drag?: string }, tag: PptTag) {
  const page = inject(injectionCurrentPage, ref(0))
  const { info, update } = useDynamicSlideInfo(computed(() => unref(page)))

  const editing = ref(false)
  const saving = ref(false)
  const draft = ref('')
  const error = ref('')

  const current = (): Located | string => {
    const content = info.value?.content
    if (content == null) return 'スライドの情報をまだ読み込めていません。少し待ってから開き直してください'
    return locate(content, tag, props.drag ?? '')
  }

  async function open() {
    error.value = ''
    for (let i = 0; i < 30 && info.value == null; i++) await new Promise((r) => setTimeout(r, 100))
    const loc = current()
    if (typeof loc === 'string') {
      error.value = loc
      draft.value = ''
    }
    else {
      draft.value = loc.inner.trim()
    }
    editing.value = true
  }

  async function save() {
    const content = info.value?.content
    const loc = current()
    if (content == null || typeof loc === 'string') {
      error.value = typeof loc === 'string' ? loc : '保存先が読めません'
      return
    }
    const next = replaceInner(content, loc, tag, draft.value)
    if (next === content) {
      editing.value = false
      return
    }
    saving.value = true
    try {
      await update({ content: next })
      editing.value = false
    }
    catch (e) {
      error.value = `保存に失敗しました: ${e instanceof Error ? e.message : String(e)}`
    }
    finally {
      saving.value = false
    }
  }

  function cancel() {
    editing.value = false
    error.value = ''
  }

  return { editing, saving, draft, error, open, save, cancel }
}
