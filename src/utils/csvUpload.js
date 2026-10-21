export const COUNTRIES = ['IND', 'BRA', 'USA', 'GBR', 'NGA']
export const SERIES = [
  ['tokens', 'total'],
  ['token_access', 'open_source'],
  ['token_access', 'proprietary'],
  ['model_region_of_origin', 'EAS'],
  ['model_region_of_origin', 'ECS'],
  ['model_region_of_origin', 'LCN'],
  ['model_region_of_origin', 'MEA'],
  ['model_region_of_origin', 'NAC'],
  ['model_region_of_origin', 'SAS'],
  ['model_region_of_origin', 'SSF'],
]
// Quorum Check grid: one answer per country and indicator; it applies to every month and segment
export const INDICATORS = [...new Set(SERIES.map(([i]) => i))]
export const MONTH_COUNT = 4
export const CELL_COUNT = COUNTRIES.length * MONTH_COUNT * SERIES.length
export const HEADERS = ['month_date', 'country_iso3', 'indicator', 'segment', 'value']
export const LONG_MAX = 9223372036854775807n

const MONTH_RE = /^\d{4}-(0[1-9]|1[0-2])$/
const VALUE_RE = /^[0-9]+$/
const COUNTRY_SET = new Set(COUNTRIES)
const INDICATOR_SET = new Set(INDICATORS)
export const SEGMENTS_BY_INDICATOR = Object.fromEntries(INDICATORS.map((i) => [i, SERIES.filter(([x]) => x === i).map(([, s]) => s)]))

export function epochMonths(epoch) {
  if (!epoch?.startDate) return null
  const m = /^(\d{4})-(\d{2})/.exec(String(epoch.startDate))
  if (!m) return null
  const year = Number(m[1])
  const month0 = Number(m[2]) - 1
  return Array.from({ length: MONTH_COUNT }, (_, i) => {
    const total = month0 + i
    const y = year + Math.floor(total / 12)
    return `${y}-${String((total % 12) + 1).padStart(2, '0')}`
  })
}

export function isCsvFile(file) {
  if (!file) return false
  const name = (file.name || '').toLowerCase()
  if (!name.endsWith('.csv')) return false
  const type = (file.type || '').toLowerCase()
  if (!type) return true
  return (
    type === 'text/csv' ||
    type === 'application/csv' ||
    type === 'text/plain' ||
    type === 'application/vnd.ms-excel'
  )
}

// One CSV line; double-quoted fields may hold commas ("12,480") and "" for a quote
export function splitCsvLine(line) {
  const cols = []
  let cur = ''
  let quoted = false
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]
    if (quoted) {
      if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++ }
      else if (ch === '"') quoted = false
      else cur += ch
    } else if (ch === '"') quoted = true
    else if (ch === ',') { cols.push(cur.trim()); cur = '' }
    else cur += ch
  }
  cols.push(cur.trim())
  return cols
}

export function rewriteSampleMonths(text, months) {
  const lines = String(text).replace(/^﻿/, '').split(/\r?\n/)
  const order = []
  return lines
    .map((line, i) => {
      if (i === 0 || line.trim() === '') return line
      const comma = line.indexOf(',')
      const csvMonth = line.slice(0, comma)
      if (!order.includes(csvMonth)) order.push(csvMonth)
      const target = months[order.indexOf(csvMonth)]
      return target ? `${target}${line.slice(comma)}` : line
    })
    .join('\n')
}

export function cellKey(country, month, indicator, segment) {
  return `${country}|${month}|${indicator}|${segment}`
}

export function pairKey(country, indicator) {
  return `${country}|${indicator}`
}

export function allCells(months) {
  const cells = []
  for (const country of COUNTRIES) {
    for (const month of months) {
      for (const [indicator, segment] of SERIES) cells.push({ country, month, indicator, segment })
    }
  }
  return cells
}

/** What a value must be for a metric, given the partner's Quorum Check answers. */
export function valueRule(quorum, country, indicator) {
  const key = pairKey(country, indicator)
  if (quorum?.ignoredPairs?.has(key)) return 'ignored'
  const ticked = quorum?.answers?.get(key)
  if (ticked === true) return 'positive'
  if (ticked === false) return 'zero'
  return 'any'
}

const fileFailure = (kind, extra = {}) => ({
  ok: false, fileError: { kind, ...extra }, errors: [], rows: [], missingCells: [], dataRowCount: 0, readyCount: 0, badRowCount: 0,
})

/**
 * Validates a data CSV for an epoch. The file has one row per country, month and series (all 200).
 * `quorum` = { answers: Map("country|indicator" -> ticked), ignoredPairs: Set("country|indicator") }:
 * a ticked metric needs values above 0, an unticked one 0, and rows of metrics below the quorum are
 * accepted but dropped before submission. Without answers (another browser answered), any whole number is accepted.
 *
 * Errors carry a `kind` so the file check can group them (see fileCheck.js).
 */
export function parseAndValidateCsv(text, months, quorum = null) {
  if (!months || months.length !== MONTH_COUNT) return fileFailure('platform')
  if (text == null || String(text).trim() === '') return fileFailure('empty')

  const rawLines = String(text).replace(/^﻿/, '').split(/\r?\n/)
  while (rawLines.length > 0 && rawLines[rawLines.length - 1].trim() === '') rawLines.pop()
  if (rawLines.length === 0) return fileFailure('empty')

  const headerLine = rawLines[0]
  const header = splitCsvLine(headerLine)
  if (header.length === 1 && headerLine.includes(';')) return fileFailure('separator', { found: headerLine })
  if (header.length !== HEADERS.length || header.some((h, i) => h !== HEADERS[i])) {
    return fileFailure('header', { found: headerLine, rows: rawLines.length })
  }

  const monthSet = new Set(months)
  const errors = []
  const badRows = new Set()
  const add = (kind, rowNum, column, found, record, extra = {}) => {
    badRows.add(rowNum)
    errors.push({ kind, row: rowNum, column, found, record, ...extra })
  }

  const parsed = []
  const seen = new Map()
  const exact = new Set()
  // Rows whose key has invalid parts still stand for one cell: the wildcard (null) parts are matched later
  const wildcards = []
  let dataRowCount = 0

  for (let i = 1; i < rawLines.length; i++) {
    const rowNum = i + 1
    const line = rawLines[i]
    if (line.trim() === '') continue
    dataRowCount++
    const cols = splitCsvLine(line)
    const record = Object.fromEntries(HEADERS.map((h, j) => [h, cols[j] ?? '']))

    if (cols.length !== HEADERS.length) {
      add('columns', rowNum, null, cols.length, record, { raw: line })
      wildcards.push({ country: null, month: null, indicator: null, segment: null })
      continue
    }

    const { month_date: month, country_iso3: country, indicator, segment, value: raw } = record
    const errBefore = errors.length

    const monthOk = MONTH_RE.test(month) && monthSet.has(month)
    if (!MONTH_RE.test(month)) add('month-format', rowNum, 'month_date', month, record)
    else if (!monthSet.has(month)) add('month-foreign', rowNum, 'month_date', month, record)

    const countryOk = COUNTRY_SET.has(country)
    if (!countryOk) add('country', rowNum, 'country_iso3', country, record)

    const indicatorOk = INDICATOR_SET.has(indicator)
    const segmentOk = indicatorOk && SEGMENTS_BY_INDICATOR[indicator].includes(segment)
    if (!indicatorOk) add('indicator', rowNum, 'indicator', indicator, record)
    else if (!segmentOk) add('segment', rowNum, 'segment', segment, record)

    let value = null
    const rule = countryOk && indicatorOk ? valueRule(quorum, country, indicator) : 'any'
    if (raw === '') add('value-empty', rowNum, 'value', raw, record)
    else if (!VALUE_RE.test(raw) || BigInt(raw) > LONG_MAX) add('value', rowNum, 'value', raw, record, { rule })
    else {
      value = BigInt(raw)
      if (rule === 'positive' && value === 0n) add('value', rowNum, 'value', raw, record, { rule })
      else if (rule === 'zero' && value !== 0n) add('value-not-zero', rowNum, 'value', raw, record)
    }

    const keyOk = monthOk && countryOk && indicatorOk && segmentOk
    if (!keyOk) {
      wildcards.push({
        country: countryOk ? country : null,
        month: monthOk ? month : null,
        indicator: indicatorOk ? indicator : null,
        segment: segmentOk ? segment : null,
      })
      continue
    }

    const key = cellKey(country, month, indicator, segment)
    const original = seen.get(key)
    if (original) {
      add(original.raw === raw ? 'duplicate-same' : 'duplicate-diff', rowNum, null, raw, record, { duplicateOf: original })
      continue
    }
    seen.set(key, { row: rowNum, record, raw })
    exact.add(key)
    if (errors.length === errBefore) parsed.push({ rowNum, month, country, indicator, segment, value })
  }

  const uncovered = allCells(months).filter((c) => !exact.has(cellKey(c.country, c.month, c.indicator, c.segment)))
  for (const w of wildcards) {
    const i = uncovered.findIndex((c) => ['country', 'month', 'indicator', 'segment'].every((k) => w[k] == null || w[k] === c[k]))
    if (i >= 0) uncovered.splice(i, 1)
  }

  const ok = errors.length === 0 && uncovered.length === 0
  return {
    ok,
    fileError: null,
    errors,
    // Metrics below the quorum get no data at all
    rows: ok ? parsed.filter((r) => valueRule(quorum, r.country, r.indicator) !== 'ignored') : [],
    missingCells: uncovered,
    dataRowCount,
    readyCount: dataRowCount - badRows.size,
    badRowCount: badRows.size,
  }
}

export function formatInt(v) {
  if (v == null || v === '') return '—'
  try {
    return BigInt(v).toLocaleString()
  } catch {
    return String(v)
  }
}

/** The Quorum Check CSV: one row per country and indicator, report = 1 (will report) or 0 (will not). */
export const QUORUM_HEADERS = ['country_iso3', 'indicator', 'report']

export function quorumSampleCsv() {
  return [QUORUM_HEADERS.join(','), ...COUNTRIES.flatMap((c) => INDICATORS.map((i) => `${c},${i},1`))].join('\n') + '\n'
}

/** Parses a Quorum Check CSV into a Map("country|indicator" -> boolean), or returns { error }. */
export function parseQuorumCsv(text) {
  const lines = String(text ?? '').replace(/^﻿/, '').split(/\r?\n/).filter((l) => l.trim() !== '')
  if (lines.length === 0) return { error: 'The file is empty.' }
  const header = splitCsvLine(lines[0])
  if (header.join(',') !== QUORUM_HEADERS.join(',')) {
    return { error: `Row 1 must be the header ${QUORUM_HEADERS.join(',')}. Found: ${lines[0]}.` }
  }
  const answers = new Map()
  for (let i = 1; i < lines.length; i++) {
    const [country, indicator, report, ...rest] = splitCsvLine(lines[i])
    const where = `Row ${i + 1}`
    if (rest.length > 0) return { error: `${where} has more than 3 values.` }
    if (!COUNTRY_SET.has(country)) return { error: `${where}: ${country || '(empty)'} is not an accepted country code. Accepted: ${COUNTRIES.join(', ')}.` }
    if (!INDICATOR_SET.has(indicator)) return { error: `${where}: ${indicator || '(empty)'} is not an accepted indicator. Accepted: ${INDICATORS.join(', ')}.` }
    if (report !== '0' && report !== '1') return { error: `${where}: report must be 1 or 0. Found: ${report || '(empty)'}.` }
    const key = pairKey(country, indicator)
    if (answers.has(key)) return { error: `${where} repeats ${country}, ${indicator}.` }
    answers.set(key, report === '1')
  }
  const total = COUNTRIES.length * INDICATORS.length
  if (answers.size !== total) return { error: `The file needs one row for each country and indicator: ${total} rows. Found: ${answers.size}.` }
  return { answers }
}
