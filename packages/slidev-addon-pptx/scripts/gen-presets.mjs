// PowerPoint の図形の定義（ECMA-376 Part 1 の presetShapeDefinitions.xml）から src/shapes/presets.ts を作る。
// 使い方: node packages/slidev-addon-pptx/scripts/gen-presets.mjs <presetShapeDefinitions.xml>
// XML はリポジトリに置かない（約 540 KB）。取得元と XML の sha256 は生成物の先頭に書く。
// 残すのは描画に要る部分だけ: avLst（adj の既定）、gdLst（数式）、rect（文字の枠）、pathLst（輪郭）。接続点と調整ハンドルは捨てる。
import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DOMParser } from '@xmldom/xmldom'

/** 取得元。LibreOffice の、この XML を最後に変えたコミットに固定する */
const SOURCE_COMMIT = '500a70ba19d9c1207fd9121531950e55a70fd940'
const SOURCE_URL = `https://raw.githubusercontent.com/LibreOffice/core/${SOURCE_COMMIT}/oox/source/drawingml/customshapes/presetShapeDefinitions.xml`

/**
 * 定義のデータの補正（図形名 → 直した rect）。
 * pie: 元は l="il" t="ir" r="it" b="ib"。ECMA の定義の誤記。t と r の名前が入れ替わっていて、枠が図形の外に出る。PowerPoint の実挙動は未確認
 */
const RECT_FIXES = {
  pie: ['il', 'it', 'ir', 'ib'],
}

const src = process.argv[2]
if (!src) {
  console.error('usage: gen-presets.mjs <presetShapeDefinitions.xml>')
  process.exit(2)
}
const out = resolve(fileURLToPath(new URL('.', import.meta.url)), '../src/shapes/presets.ts')

const xml = readFileSync(src)
const sha256 = createHash('sha256').update(xml).digest('hex')
const doc = new DOMParser().parseFromString(xml.toString('utf8'), 'text/xml')
const kids = (el, name) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && (!name || n.localName === name))
// 数式は空白が 2 つ続くことがある（`*/ vc  vf 100000`）
const fmla = (s) => s.trim().split(/\s+/).join(' ')
const guides = (el) => (el ? kids(el, 'gd').map((g) => [g.getAttribute('name'), fmla(g.getAttribute('fmla'))]) : [])
const pt = (el) => [el.getAttribute('x'), el.getAttribute('y')]

const shapes = {}
for (const shape of kids(doc.documentElement)) {
  const s = { av: guides(kids(shape, 'avLst')[0]), gd: guides(kids(shape, 'gdLst')[0]), paths: [] }
  const rect = kids(shape, 'rect')[0]
  if (rect) s.rect = ['l', 't', 'r', 'b'].map((k) => rect.getAttribute(k))
  if (RECT_FIXES[shape.localName]) {
    // 元の XML が直っていたら補正は要らない。表から消すよう止める
    if (!s.rect || JSON.stringify(s.rect) === JSON.stringify(RECT_FIXES[shape.localName])) {
      throw new Error(`${shape.localName}: rect の補正が要らない（元の rect ${JSON.stringify(s.rect)}）。RECT_FIXES から消す`)
    }
    s.rect = RECT_FIXES[shape.localName]
  }
  for (const path of kids(kids(shape, 'pathLst')[0], 'path')) {
    const p = { cmds: [] }
    if (path.getAttribute('w')) p.w = Number(path.getAttribute('w'))
    if (path.getAttribute('h')) p.h = Number(path.getAttribute('h'))
    const fill = path.getAttribute('fill')
    if (fill && fill !== 'norm') p.fill = fill
    if (path.getAttribute('stroke') === 'false') p.stroke = false
    for (const c of kids(path)) {
      const pts = kids(c, 'pt').flatMap(pt)
      switch (c.localName) {
        case 'moveTo': p.cmds.push(['M', ...pts]); break
        case 'lnTo': p.cmds.push(['L', ...pts]); break
        case 'quadBezTo': p.cmds.push(['Q', ...pts]); break
        case 'cubicBezTo': p.cmds.push(['C', ...pts]); break
        case 'arcTo': p.cmds.push(['A', c.getAttribute('wR'), c.getAttribute('hR'), c.getAttribute('stAng'), c.getAttribute('swAng')]); break
        case 'close': p.cmds.push(['Z']); break
        default: throw new Error(`${shape.localName}: 未知の path 命令 ${c.localName}`)
      }
    }
    s.paths.push(p)
  }
  shapes[shape.localName] = s
}

for (const n of Object.keys(RECT_FIXES)) if (!shapes[n]) throw new Error(`RECT_FIXES の ${n} が定義に無い`)
const fixes = Object.keys(RECT_FIXES)

const names = Object.keys(shapes).sort()
const body = names.map((n) => `  ${n}: ${JSON.stringify(shapes[n])},`).join('\n')
writeFileSync(
  out,
  `// 生成物。手で直さない。scripts/gen-presets.mjs で作り直す。
// 元: ECMA-376 Part 1（Office Open XML File Formats, Fundamentals and Markup Language Reference）の
// presetShapeDefinitions.xml。PowerPoint の図形 ${names.length} 種の定義。Copyright © Ecma International.
// 補正 ${fixes.length} 件（${fixes.map((n) => `${n} の rect`).join('、')}）。理由は scripts/gen-presets.mjs の RECT_FIXES
// 取得元: ${SOURCE_URL}
// XML の sha256: ${sha256}
import type { PresetShape } from './types.ts'

export const PRESETS: Record<string, PresetShape> = {
${body}
}
`,
)
console.log(`${names.length} shapes (sha256 ${sha256}) -> ${out}`)
