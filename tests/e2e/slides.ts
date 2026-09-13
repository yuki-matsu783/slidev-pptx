// plain.md のスライド番号。デッキを直したらここだけ直す。
// 注意: @slidev/parser は `---` で始まる行を無条件に区切りにする（`dist/core.mjs` の parseSync）。
// 水平線が欲しいときは `***` か <hr> を使う。
export const S = {
  cover: 1,
  bullets: 2,
  twoCols: 3,
  table: 4,
  math: 5,
  boxes: 6,
  center: 7,
  section: 8,
  imageRight: 9,
  coverBg: 10,
  zoom: 11,
  offslide: 12,
  twoColsHeader: 13,
} as const

export const SLIDE_COUNT = 13
