// CLI（native-export.md §7）。exit code: 0 成功 / 1 書き出し失敗または OPC error（--strict なら警告も）/ 2 引数の誤り
import { readFileSync, existsSync } from 'node:fs'
import { exportPptx, isValidRange, summaryLines } from './export.ts'
import { check } from './opc/check.ts'

const USAGE = `slidev-pptx export [entry=slides.md]
  --output, -o <path>     既定 ./slides-export.pptx
  --range <1-3,5>         Slidev の range と同じ綴り
  --theme <name>
  --lang <tag>            run の lang の既定。既定 ja-JP
  --wait <ms>             描画後の追加待機。既定 0
  --timeout <ms>          ページ読み込みの上限。既定 30000
  --report <path>         記録の JSON。既定 <output>.report.json
  --no-check              OPC 整合チェックを飛ばす
  --strict                警告が 1 つでもあれば exit 1
  --keep-server           失敗時にサーバとブラウザを閉じない
slidev-pptx check <file.pptx>`

interface Parsed {
  command: string
  positional: string[]
  flags: Record<string, string | boolean>
}

const KNOWN_FLAGS: Record<string, 'string' | 'boolean'> = {
  output: 'string',
  o: 'string',
  range: 'string',
  theme: 'string',
  lang: 'string',
  wait: 'string',
  timeout: 'string',
  report: 'string',
  'no-check': 'boolean',
  strict: 'boolean',
  'keep-server': 'boolean',
  help: 'boolean',
  h: 'boolean',
}

export function parseArgs(argv: string[]): Parsed | { error: string } {
  // `slidev-pptx --help` のように先頭がフラグならサブコマンド無し
  const command = argv[0] && !argv[0].startsWith('-') ? argv[0] : ''
  const rest = command ? argv.slice(1) : argv
  const positional: string[] = []
  const flags: Record<string, string | boolean> = {}
  for (let i = 0; i < rest.length; i++) {
    const a = rest[i]
    if (a.startsWith('-')) {
      const name = a.replace(/^-{1,2}/, '')
      const [key, inline] = name.includes('=') ? name.split(/=(.*)/s) : [name, undefined]
      const kind = KNOWN_FLAGS[key]
      if (!kind) return { error: `未知のフラグ: ${a}` }
      if (kind === 'boolean') flags[key] = true
      else {
        const v = inline ?? rest[++i]
        if (v === undefined) return { error: `${a} には値が要る` }
        flags[key] = v
      }
    } else positional.push(a)
  }
  return { command, positional, flags }
}

export async function main(argv: string[]): Promise<number> {
  const parsed = parseArgs(argv)
  if ('error' in parsed) {
    console.error(parsed.error)
    console.error(USAGE)
    return 2
  }
  const { command, positional, flags } = parsed
  if (flags.help || flags.h) {
    console.log(USAGE)
    return 0
  }
  if (command === 'check') {
    const file = positional[0]
    if (!file) {
      console.error('check には PPTX のパスが要る')
      return 2
    }
    if (!existsSync(file)) {
      console.error(`ファイルが無い: ${file}`)
      return 1
    }
    const results = await check(readFileSync(file))
    for (const r of results) console.log(`${r.level.padEnd(5)} ${r.rule.padEnd(4)} ${r.part}: ${r.message}`)
    const errors = results.filter((r) => r.level === 'error').length
    console.log(`${file}: ${errors} errors, ${results.length - errors} warnings`)
    return errors ? 1 : 0
  }
  if (command === 'export') {
    const entry = positional[0] ?? 'slides.md'
    for (const k of ['wait', 'timeout'] as const) {
      if (flags[k] !== undefined && !Number.isFinite(Number(flags[k]))) {
        console.error(`--${k} には数値が要る: ${flags[k]}`)
        return 2
      }
    }
    if (flags.range !== undefined && !isValidRange(String(flags.range))) {
      console.error(`--range の綴りが不正: ${flags.range}（例: 1-3,5）`)
      return 2
    }
    if (!existsSync(entry)) {
      console.error(`entry が無い: ${entry}`)
      return 1
    }
    try {
      const report = await exportPptx({
        entry,
        output: (flags.output ?? flags.o) as string | undefined,
        range: flags.range as string | undefined,
        theme: flags.theme as string | undefined,
        lang: flags.lang as string | undefined,
        wait: flags.wait ? Number(flags.wait) : undefined,
        timeout: flags.timeout ? Number(flags.timeout) : undefined,
        report: flags.report as string | undefined,
        check: !flags['no-check'],
        strict: !!flags.strict,
        keepServer: !!flags['keep-server'],
        log: (line) => console.log(line),
      })
      const errors = report.check.filter((c) => c.level === 'error').length
      if (errors) {
        console.error(`OPC 整合チェックで ${errors} 件の error。ファイルは残してある: ${report.output}`)
        return 1
      }
      if (flags.strict && report.warnings.length) {
        console.error(`--strict: 警告が ${report.warnings.length} 件`)
        return 1
      }
      return 0
    } catch (e) {
      console.error(e instanceof Error ? e.stack ?? e.message : String(e))
      return 1
    }
  }
  console.error(command ? `未知のサブコマンド: ${command}` : USAGE)
  return 2
}

export { summaryLines }
