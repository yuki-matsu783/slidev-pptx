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

function expectFinite(name: string, w: number, h: number, adj?: Record<string, number>) {
  const g = evalPreset(name, w, h, adj)
  expect(g.paths.length).toBeGreaterThan(0)
  for (const p of g.paths) {
    expect(p.d).not.toMatch(/NaN|Infinity/)
    expect(p.d).toMatch(/^M/)
  }
  for (const v of Object.values(g.textRect)) expect(Number.isFinite(v)).toBe(true)
}

describe('shapes/geometry: 名前と引数', () => {
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
  it('有限でない大きさは例外', () => {
    expect(() => evalPreset('rect', Number.NaN, 50)).toThrow(/有限/)
    expect(() => evalPreset('rect', 50, Number.POSITIVE_INFINITY)).toThrow(/有限/)
  })
})

describe('shapes/geometry: 187 種すべて', () => {
  const sizes: [number, number][] = [[200, 100], [100, 200], [150, 150], [0, 0], [0, 100]]
  for (const name of PRESET_NAMES) {
    it(`${name} は 3 通りの縦横比と大きさ 0 で有限の値になる`, () => {
      for (const [w, h] of sizes) expectFinite(name, w, h)
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

  it('at2 は at2 x y = atan2(y, x): flowChartMagneticTape の最後の円弧は枠の対角の向きで楕円に当たる', () => {
    // ang1 = at2 w h = atan2(100, 200) = 26.565°。見た目の角度 θ の媒介変数 t = atan2(100·sinθ, 50·cosθ) = 45°
    // → 終点 (100 + 100·cos45°, 50 + 50·sin45°) = (170.711, 85.355)。ib = vc + sin hd2 45° = 85.355
    const g = evalPreset('flowChartMagneticTape', 200, 100)
    expect(g.paths[0].d).toBe(
      'M100 100 A100 50 0 0 1 0 50 A100 50 0 0 1 100 0 A100 50 0 0 1 200 50 A100 50 0 0 1 170.711 85.355 L200 85.355 L200 100 Z',
    )
  })

  it('tan は tan x y = x·tan(y): swooshArrow の矢じり', () => {
    // alfa = */ cd4 1 14 = 90°/14、tan(alfa) = 0.112673。ssd8 = 12.5、ad1 = h·a1/100000 = 25、ad2 = ss·a2/100000 = 16.667
    // xB = r − ad2 = 183.333、yB = ssd8 = 12.5、dx0 = tan ssd8 alfa = 1.408 → xC = 181.925
    // dx1 = tan ad1 alfa = 2.817 → xF = 186.150、xE = xF + dx0 = 187.558、yF = 37.5、yE = 50、yD = +- t dy22 dy3 = 25 − 5 = 20
    // xP1 = wd6 = 33.333、yP1 = hd6 + hd6 = 33.333、xP2 = wd4 = 50、yP2 = yF + hd6/2 = 45.833
    const cmds = parse(evalPreset('swooshArrow', 200, 100).paths[0].d)
    expect(cmds.map((c) => c.op).join('')).toBe('MQLLLLQZ')
    expectPoints(cmds[1].args, [33.333, 33.333, 183.333, 12.5])
    expectPoints(cmds[2].args, [181.925, 0])
    expectPoints(cmds[3].args, [200, 20])
    expectPoints(cmds[4].args, [187.558, 50])
    expectPoints(cmds[5].args, [186.15, 37.5])
    expectPoints(cmds[6].args, [50, 45.833, 0, 100])
  })

  it('sqrt と +/ と Q: teardrop の既定', () => {
    // r2 = sqrt 2、tw = wd2·r2 = 141.421、th = hd2·r2 = 70.711、a = 100000 → sw = tw、sh = th
    // dx1 = cos sw 45° = 100、dy1 = sin sh 45° = 50 → x1 = hc + 100 = 200、y1 = vc − 50 = 0
    // x2 = +/ hc x1 2 = (100 + 200)/2 = 150、y2 = +/ vc y1 2 = 25
    expect(evalPreset('teardrop', 200, 100).paths[0].d).toBe(
      'M0 50 A100 50 0 0 1 100 0 Q150 0 200 0 Q200 25 200 50 A100 50 0 0 1 100 100 A100 50 0 0 1 0 50 Z',
    )
  })

  it('max と min: wave の adj2 = 5000 の文字の枠', () => {
    // dx1 = w·a2/100000 = 10、of2 = 20 → dx2 = ?: of2 0 of2 = 0、dx5 = 20 → x2 = 0、x5 = 180、x6 = 20、x10 = 200
    // dx3 = (dx2 + x5)/3 = 60 → x3 = 60、x4 = (x3 + x5)/2 = 120、x7 = x6 + dx3 = 80、x8 = (x7 + x10)/2 = 140
    // y1 = h·a1/100000 = 12.5、dy2 = y1·10/3 = 41.667 → y2 = −29.167、y3 = 54.167、y4 = 87.5、y5 = 45.833、y6 = 129.167
    // il = max x2 x6 = 20、ir = min x5 x10 = 180、it = h·a1/50000 = 25、ib = 75
    const g = evalPreset('wave', 200, 100, { adj2: 5000 })
    expect(g.paths[0].d).toBe('M0 12.5 C60 -29.167 120 54.167 180 12.5 L200 87.5 C140 129.167 80 45.833 20 87.5 Z')
    expect(g.textRect).toEqual({ l: 20, t: 25, r: 180, b: 75 })
  })

  it('abs: wedgeRectCallout の尾が上（adj2 = −62500）', () => {
    // dxPos = w·(−20833)/100000 = −41.666、dyPos = −62.5 → xPos = 58.334、yPos = −12.5
    // dq = dxPos·h/w = −20.833、dz = abs dyPos − abs dq = 41.667 > 0 → 尾は上の辺（xt = xPos、yt = yPos）
    // dxPos < 0 → x1 = w·2/12 = 33.333、x2 = w·5/12 = 83.333。dyPos < 0 → y1 = h·2/12 = 16.667、y2 = 41.667
    expect(evalPreset('wedgeRectCallout', 200, 100, { adj2: -62500 }).paths[0].d).toBe(
      'M0 0 L33.333 0 L58.334 -12.5 L83.333 0 L200 0 L200 16.667 L200 16.667 L200 41.667 L200 100 L83.333 100 L33.333 100 L33.333 100 L0 100 L0 41.667 L0 16.667 L0 16.667 Z',
    )
  })

  it('?: は x > 0 のときだけ y: chevron 100×200 は dx = 0 で文字の枠が枠全体', () => {
    // ss = 100、a = 50000 → x1 = 50、x2 = r − x1 = 50、dx = x2 − x1 = 0 → il = ?: dx x1 l = l、ir = ?: dx x2 r = r
    const g = evalPreset('chevron', 100, 200)
    expect(g.paths[0].d).toBe('M0 0 L50 0 L100 100 L50 200 L0 200 L50 100 Z')
    expect(g.textRect).toEqual({ l: 0, t: 0, r: 100, b: 200 })
  })

  it('ssdN は ss/N: stripedRightArrow 200×100（ss = 100）の縞', () => {
    // ssd32 = 3.125、ssd16 = 6.25、ssd8 = 12.5、x4 = ss·5/32 = 15.625。dy1 = h·a1/200000 = 25、dx5 = ss·a2/100000 = 50 → x5 = 150
    // dx6 = dy1·dx5/hd2 = 25 → x6 = 175
    const g = evalPreset('stripedRightArrow', 200, 100)
    expect(g.paths[0].d).toBe(
      'M0 25 L3.125 25 L3.125 75 L0 75 Z M6.25 25 L12.5 25 L12.5 75 L6.25 75 Z M15.625 25 L150 25 L150 0 L200 50 L150 100 L150 75 L15.625 75 Z',
    )
    expect(g.textRect).toEqual({ l: 15.625, t: 25, r: 175, b: 75 })
  })

  it('mod は sqrt(x² + y² + z²): cornerTabs 300×400 の角の三角は対角 500 の 1/20', () => {
    // md = mod w h 0 = 500、dx = md/20 = 25、x1 = r − dx = 275、y1 = b − dx = 375
    const g = evalPreset('cornerTabs', 300, 400)
    expect(g.paths.map((p) => p.d)).toEqual(['M0 0 L25 0 L0 25 Z', 'M0 375 L25 400 L0 400 Z', 'M275 0 L300 0 L300 25 Z', 'M300 375 L300 400 L275 400 Z'])
    expect(g.textRect).toEqual({ l: 25, t: 25, r: 275, b: 375 })
  })

  it('負の swAng は sweep 0: flowChartOnlineStorage（path は 6×6）', () => {
    // 60×60 なので倍率 10。moveTo 1,0 → (10,0)、lnTo 6,0 → (60,0)
    // arcTo wR 1 hR 3 3cd4 −cd2: 中心 = (60,0) − (10·cos270°, 30·sin270°) = (60,30)、終点 90° = (60,60)
    // lnTo 1,6 → (10,60)、arcTo 1 3 cd4 cd2: 中心 (10,30)、終点 270° = (10,0)
    expect(evalPreset('flowChartOnlineStorage', 60, 60).paths[0].d).toBe('M10 0 L60 0 A10 30 0 0 0 60 60 L10 60 A10 30 0 0 1 10 0 Z')
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
})

describe('shapes/geometry: 回帰（レビューの指摘）', () => {
  it('半径 0 の円弧でも 90 度の倍数の向きが狂わない: leftBracket の adj = 0 は角の丸みの無い [', () => {
    // y1 = ss·a/100000 = 0。arcTo w y1 cd4 cd4: hR = 0、中心 = (160,100) − (160·cos90°, 0) = (160,100)、終点 180° = (0,100)
    // lnTo l y1 = (0,0)、arcTo w y1 cd2 cd4: 中心 (160,0)、終点 270° = (160,0)。Math.sin(π) ≠ 0 のままだと斜線になっていた
    expect(evalPreset('leftBracket', 160, 100, { adj: 0 }).paths[1].d).toBe('M160 100 A160 0 0 0 1 0 100 L0 0 A160 0 0 0 1 160 0')
  })

  it('幅 0 の楕円でも軸の向きの点に届く: donut 100×200 の adj = 50000 の内側', () => {
    // dr = ss·a/100000 = 50 → iwd2 = wd2 − dr = 0、ihd2 = hd2 − dr = 50。moveTo dr vc = (50,100)
    // arcTo 0 50 cd2 −cd4: 中心 (50,100)、終点 90° = (50,150) → 0° = (50,100) → 270° = (50,50) → 180° = (50,100)
    const d = evalPreset('donut', 100, 200, { adj: 50000 }).paths[0].d
    expect(d.slice(d.lastIndexOf('M'))).toBe('M50 100 A0 50 0 0 0 50 150 A0 50 0 0 0 50 100 A0 50 0 0 0 50 50 A0 50 0 0 0 50 100 Z')
  })

  it('範囲内の adj で 0 割り・負の sqrt になっても NaN にしない（値は近似）', () => {
    // circularArrow 系 adj5 = 0: +/ q11 q10 q4 の q4 = 0。pentagon hf = 0: */ y1 dx2 dx1 の dx1 = 0
    // leftUpArrow / leftRightUpArrow / quadArrow adj2 = 0: il の分母 0。curved*Arrow: sqrt の中が負
    expectFinite('circularArrow', 160, 100, { adj5: 0 })
    expectFinite('leftCircularArrow', 160, 100, { adj5: 0 })
    expectFinite('leftRightCircularArrow', 160, 100, { adj5: 0 })
    expectFinite('pentagon', 160, 100, { hf: 0 })
    expectFinite('leftUpArrow', 160, 100, { adj2: 0 })
    expectFinite('leftRightUpArrow', 160, 100, { adj2: 0 })
    expectFinite('quadArrow', 160, 100, { adj2: 0 })
    expectFinite('curvedUpArrow', 100, 100, { adj1: 51000 })
    expectFinite('curvedDownArrow', 160, 100, { adj1: 100000 })
  })

  it('180 度を超える円弧は割る: arc の adj1 = 0, adj2 = 21599999 が消えない', () => {
    // stAng = 0 → 始点 (hc + wd2, vc) = (200,50)。swAng = 21599999（360° 未満）を 2 つに割る → ほぼ (0,50) を経て (200,50) へ
    // 1 つの A だと始点と終点が丸めで重なり、SVG は円弧を描かない
    expect(evalPreset('arc', 200, 100, { adj1: 0, adj2: 21599999 }).paths[1].d).toBe('M200 50 A100 50 0 0 1 0 50 A100 50 0 0 1 200 50')
  })

  it('arc の adj1 = adj2 は 1 周（swAng = ?: sw11 sw11 sw12 で sw11 = 0 は 0 より大きくない）', () => {
    expect(evalPreset('arc', 200, 100, { adj1: 0, adj2: 0 }).paths[1].d).toBe('M200 50 A100 50 0 0 1 0 50 A100 50 0 0 1 200 50')
  })
})
