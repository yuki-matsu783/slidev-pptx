// PowerPoint の図形の定義（ECMA-376 Part 1 の presetShapeDefinitions.xml）から src/shapes/presets.ts を作る。
// 使い方: node packages/slidev-addon-pptx/scripts/gen-presets.mjs <presetShapeDefinitions.xml>
// XML はリポジトリに置かない（約 540 KB）。取得元は生成物の先頭に書く。
// 残すのは描画に要る部分だけ: avLst（adj の既定）、gdLst（数式）、rect（文字の枠）、pathLst（輪郭）。接続点と調整ハンドルは捨てる。
import { readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DOMParser } from '@xmldom/xmldom'

const src = process.argv[2]
if (!src) {
  console.error('usage: gen-presets.mjs <presetShapeDefinitions.xml>')
  process.exit(2)
}
const out = resolve(fileURLToPath(new URL('.', import.meta.url)), '../src/shapes/presets.ts')

const doc = new DOMParser().parseFromString(readFileSync(src, 'utf8'), 'text/xml')
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

const names = Object.keys(shapes).sort()
const body = names.map((n) => `  ${n}: ${JSON.stringify(shapes[n])},`).join('\n')
writeFileSync(
  out,
  `// 生成物。手で直さない。scripts/gen-presets.mjs で作り直す。
// 元: ECMA-376 Part 1 の presetShapeDefinitions.xml（PowerPoint の図形 ${names.length} 種の定義）。
// 取得元: https://raw.githubusercontent.com/LibreOffice/core/master/oox/source/drawingml/customshapes/presetShapeDefinitions.xml
import type { PresetShape } from './types.ts'

export const PRESETS: Record<string, PresetShape> = {
${body}
}
`,
)
console.log(`${names.length} shapes -> ${out}`)
