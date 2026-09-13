// 図形の評価器（src/shapes/geometry.ts）。答えは presetShapeDefinitions.xml の数式から手で導いたもの
import { describe, expect, it } from 'vitest'
import { PRESET_NAMES, evalPreset, isPreset } from '../../packages/slidev-addon-pptx/src/shapes/geometry'

type Cmd = { op: string; args: number[] }

/** d を命令ごとに割る（A は rx ry rot large sweep x y の 7 個） */
function parse(d: string): Cmd[] {
  const out: Cmd[] = []
  for (const tok of d.match(/[MLQCAZ]|-?\d+(\.\d+)?(e-?\d+)?/g) ?? []) {
    if (/[MLQCAZ]/.test(tok)) out.push({ op: tok, args: [] })
    else out[out.length - 1].args.push(Number(tok))
  }
  return out
}

function expectPoints(actual: number[], expected: number[]) {
  expect(actual).toHaveLength(expected.length)
  actual.forEach((v, i) => expect(v).toBeCloseTo(expected[i], 2))
}

describe('shapes/geometry: 名前', () => {
  it('187 種で、isPreset は定義にある名前だけ真', () => {
    expect(PRESET_NAMES).toHaveLength(187)
    expect(isPreset('rect')).toBe(true)
    expect(isPreset('foldedCorner')).toBe(true)
    expect(isPreset('bentConnector3')).toBe(true)
    expect(isPreset('folderCorner')).toBe(false)
    expect(isPreset('toString')).toBe(false)
  })
  it('未知の名前は例外', () => {
    expect(() => evalPreset('nope', 100, 100)).toThrow(/nope/)
  })
})

describe('shapes/geometry: 187 種すべて', () => {
  const sizes: [number, number][] = [[200, 100], [100, 200], [150, 150]]
  for (const name of PRESET_NAMES) {
    it(`${name} は 3 通りの縦横比で有限の値になる`, () => {
      for (const [w, h] of sizes) {
        const g = evalPreset(name, w, h)
        expect(g.paths.length).toBeGreaterThan(0)
        for (const p of g.paths) {
          expect(p.d).not.toMatch(/NaN|Infinity/)
          expect(p.d).toMatch(/^M/)
        }
        for (const v of Object.values(g.textRect)) expect(Number.isFinite(v)).toBe(true)
      }
    })
  }
})

describe('shapes/geometry: 座標', () => {
  it('rect は枠そのもの', () => {
    // path: l,t → r,t → r,b → l,b。rect l t r b
    const g = evalPreset('rect', 200, 100)
    expect(g.paths).toEqual([{ d: 'M0 0 L200 0 L200 100 L0 100 Z', fill: 'norm', stroke: true }])
    expect(g.textRect).toEqual({ l: 0, t: 0, r: 200, b: 100 })
  })

  it('triangle の頂点 x は w·adj/100000、文字の枠は x1=w·a/200000 から x1+wd2、vc から b', () => {
    // gd: x2 = */ w a 100000、x1 = */ w a 200000、x3 = +- x1 wd2 0
    const g = evalPreset('triangle', 200, 100)
    expect(g.paths[0].d).toBe('M0 100 L100 0 L200 100 Z')
    expect(g.textRect).toEqual({ l: 50, t: 50, r: 150, b: 100 })
    expect(evalPreset('triangle', 200, 100, { adj: 25000 }).paths[0].d).toBe('M0 100 L50 0 L200 100 Z')
  })

  it('rightArrow の既定（200×100）', () => {
    // dx1 = ss·a2/100000 = 50 → x1 = 150。dy1 = h·a1/200000 = 25 → y1 = 25, y2 = 75。dx2 = y1·dx1/hd2 = 25 → x2 = 175
    const g = evalPreset('rightArrow', 200, 100)
    expect(g.paths[0].d).toBe('M0 25 L150 25 L150 0 L200 50 L150 100 L150 75 L0 75 Z')
    expect(g.textRect).toEqual({ l: 0, t: 25, r: 175, b: 75 })
  })

  it('roundRect の adj は ss 比の半径で、角の円弧は中心 (x1, x1) の 4 分の 1 円', () => {
    // x1 = */ ss a 100000 = 100·25000/100000 = 25。il = */ x1 29289 100000
    const g = evalPreset('roundRect', 200, 100, { adj: 25000 })
    expect(g.paths[0].d).toBe(
      'M0 25 A25 25 0 0 1 25 0 L175 0 A25 25 0 0 1 200 25 L200 75 A25 25 0 0 1 175 100 L25 100 A25 25 0 0 1 0 75 Z',
    )
    expect(g.textRect.l).toBeCloseTo(25 * 0.29289, 6)
    expect(g.textRect.b).toBeCloseTo(100 - 25 * 0.29289, 6)
  })

  it('adj は pin で定義の範囲に収まる', () => {
    // rightArrow: maxAdj2 = */ 100000 w ss = 200000 → a2 = 200000 → dx1 = ss·a2/100000 = 200 → x1 = 0
    const g = evalPreset('rightArrow', 200, 100, { adj2: 300000 })
    expect(parse(g.paths[0].d)[1]).toEqual({ op: 'L', args: [0, 25] })
  })

  it('定義に無い adj のキーは無視する', () => {
    expect(evalPreset('triangle', 200, 100, { adj9: 1 })).toEqual(evalPreset('triangle', 200, 100))
  })

  it('ellipse は l,vc から始まる 4 つの円弧で、中心は (hc, vc)', () => {
    // arcTo wd2 hd2 cd2 cd4: 始点 (0,50) − (100·cos180°, 50·sin180°) = 中心 (100,50)、終点 = 中心 + (100·cos270°, 50·sin270°) = (100,0)
    const g = evalPreset('ellipse', 200, 100)
    expect(g.paths[0].d).toBe('M0 50 A100 50 0 0 1 100 0 A100 50 0 0 1 200 50 A100 50 0 0 1 100 100 A100 50 0 0 1 0 50 Z')
    // idx = cos wd2 45°、idy = sin hd2 45°
    expect(g.textRect.l).toBeCloseTo(100 - 100 * Math.SQRT1_2, 6)
    expect(g.textRect.t).toBeCloseTo(50 - 50 * Math.SQRT1_2, 6)
  })

  it('arcTo の角度は見た目の角度: arc の始点（数式）と円弧の中心（換算）が合う', () => {
    // arc adj1 = 315°: wt1 = sin wd2 315° = −70.711、ht1 = cos hd2 315° = 35.355、
    // dx1 = cat2 wd2 ht1 wt1 = 100·cos(atan2(−70.711, 35.355)) = 44.721、dy1 = sat2 = 50·sin(同) = −44.721 → 始点 (144.721, 5.279)
    // swAng = 0 − 315° + 360° = 45° → 終点は 0° の (200, 50)、large-arc 0、sweep 1
    const g = evalPreset('arc', 200, 100, { adj1: 18900000, adj2: 0 })
    const cmds = parse(g.paths[1].d)
    expect(cmds.map((c) => c.op)).toEqual(['M', 'A'])
    expectPoints(cmds[0].args, [144.721, 5.279])
    expectPoints(cmds[1].args, [100, 50, 0, 0, 1, 200, 50])
    expect(g.paths[0]).toMatchObject({ fill: 'norm', stroke: false })
    expect(g.paths[1]).toMatchObject({ fill: 'none', stroke: true })
  })

  it('360 度の円弧は 2 つの半円に割る（actionButtonInformation の円）', () => {
    // arcTo r r 3cd4 21600000: 始点は円の頂上 → 半周で真下 (x, y+2r) → もう半周で始点に戻る
    const g = evalPreset('actionButtonInformation', 200, 100)
    let found = 0
    for (const p of g.paths) {
      const cmds = parse(p.d)
      for (let i = 0; i + 2 < cmds.length; i++) {
        const [m, a1, a2] = cmds.slice(i, i + 3)
        if (m.op !== 'M' || a1.op !== 'A' || a2.op !== 'A') continue
        const [x, y] = m.args
        const r = a1.args[0]
        expectPoints(a1.args, [r, r, 0, 0, 1, x, y + 2 * r])
        expectPoints(a2.args, [r, r, 0, 0, 1, x, y])
        found++
      }
    }
    expect(found).toBeGreaterThan(0)
  })

  it('star5 は頂点 10 個の閉じた折れ線', () => {
    const cmds = parse(evalPreset('star5', 200, 200).paths[0].d)
    expect(cmds.map((c) => c.op).join('')).toBe('M' + 'L'.repeat(9) + 'Z')
    // 頂上の頂点は (hc, t)
    expect(cmds[2].args).toEqual([100, 0])
  })

  it('path に w h があれば、その座標系から枠へ拡大する（cloud は 43200×43200）', () => {
    // moveTo 3900,14370 → (3900·200/43200, 14370·100/43200)。arcTo wR 6753 hR 9190 → (6753·200/43200, 9190·100/43200)
    const cmds = parse(evalPreset('cloud', 200, 100).paths[0].d)
    expectPoints(cmds[0].args, [(3900 * 200) / 43200, (14370 * 100) / 43200])
    expect(cmds[1].op).toBe('A')
    expect(cmds[1].args[0]).toBeCloseTo((6753 * 200) / 43200, 2)
    expect(cmds[1].args[1]).toBeCloseTo((9190 * 100) / 43200, 2)
    // flowChartProcess は path w=1 h=1 の正方形 → 枠いっぱい
    expect(evalPreset('flowChartProcess', 200, 100).paths[0].d).toBe('M0 0 L200 0 L200 100 L0 100 Z')
  })

  it('rect の無い図形の文字の枠は枠全体、fill="none" と stroke="false" を写す', () => {
    expect(evalPreset('line', 200, 100).textRect).toEqual({ l: 0, t: 0, r: 200, b: 100 })
    // actionButtonHome: 1 つ目は stroke="false"、2 つ目は fill="darkenLess"、3 つ目は fill="darken"、4 つ目は fill="none"
    const g = evalPreset('actionButtonHome', 100, 100)
    expect(g.paths.map((p) => [p.fill, p.stroke])).toEqual([
      ['norm', false], ['darkenLess', false], ['darken', false], ['none', true], ['none', true],
    ])
  })

  it('大きさ 0 でも NaN にならない', () => {
    for (const name of ['rightArrow', 'roundRect', 'star5']) {
      const g = evalPreset(name, 0, 0)
      for (const p of g.paths) expect(p.d).not.toMatch(/NaN|Infinity/)
    }
  })
})
