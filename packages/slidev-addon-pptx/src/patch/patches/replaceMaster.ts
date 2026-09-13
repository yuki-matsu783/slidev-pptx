// 後処理 6: 将来のマスター差し替えの継ぎ目（native-export.md §5.4）。今回は no-op。
// 入力（MasterSwap）を渡す口だけ用意し、渡されなければ何もしない。
// 前提にする変換: 1〜5 の後、7（rebuildContentTypes）の前
import type { Patch } from '../index.ts'

export interface MasterSwap {
  /** 既存 .pptx（テンプレート） */
  template: Buffer
  /** Slidev のレイアウト名 → テンプレートの <p:cSld name> */
  layoutMap: Record<string, string>
  /** Slidev のレイアウト名 → placeholder 名 → テンプレート側の <p:ph idx type> */
  placeholderMap: Record<string, Record<'title' | 'body' | 'body2', { idx: number; type: string }>>
}

let pending: MasterSwap | undefined

/** 次の postProcess で使う差し替えを登録する（今は誰も呼ばない） */
export function setMasterSwap(swap: MasterSwap | undefined): void {
  pending = swap
}

export const replaceMaster: Patch = {
  name: 'replaceMaster',
  async run() {
    if (!pending) return
    // 手順（調査 §2 の見立て）: テンプレートの master / layouts / theme / media を置く → slideN.xml.rels の slideLayout
    // 関係（findRelByType で引く）の Target を layoutMap で付け替える → <p:ph idx type> を placeholderMap で付け替える。
    // 実装は既存テンプレートへの差し替えを扱うフェーズで。
    throw new Error('replaceMaster: マスター差し替えはまだ実装されていない')
  },
}
