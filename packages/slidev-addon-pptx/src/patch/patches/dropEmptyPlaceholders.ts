// 後処理 2: スライドの <p:ph> を持つ <p:sp> のうち、<a:t> に文字が無いものを消す（native-export.md §3.2）。
// masterName 付きのスライドには、埋めていない placeholder が空の図形として必ず足されるため。レイアウトは触らない。
// 前提にする変換: 1（id は振り直し済み。消しても id は一意のまま）
import { xmlPatch } from '../index.ts'
import { elements } from '../zip.ts'

export const dropEmptyPlaceholders = xmlPatch('dropEmptyPlaceholders', /^ppt\/slides\/slide\d+\.xml$/, (doc) => {
  for (const sp of elements(doc, 'p', 'sp')) {
    if (elements(sp, 'p', 'ph').length === 0) continue
    const text = elements(sp, 'a', 't')
      .map((t) => t.textContent ?? '')
      .join('')
    if (text.length === 0) sp.parentNode?.removeChild(sp)
  }
})
