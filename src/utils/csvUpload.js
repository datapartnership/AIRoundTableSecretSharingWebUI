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
export const MONTH_COUNT = 4
export const CELL_COUNT = COUNTRIES.length * MONTH_COUNT * SERIES.length
export const HEADERS = ['month_date', 'country_iso3', 'indicator', 'segment', 'value']
export const ERROR_DISPLAY_CAP = 20
export const LONG_MAX = 9223372036854775807n

const MONTH_RE = /^\d{4}-(0[1-9]|1[0-2])$/
const VALUE_RE = /^[0-9]+$/
const SERIES_SET = new Set(SERIES.map(([i, s]) => `${i}|${s}`))
const COUNTRY_SET = new Set(COUNTRIES)
const INDICATOR_SET = new Set(SERIES.map(([i]) => i))

export function epochMonths(epoch) {
  if (!epoch?.startDate) return null
  const d0 = new Date(epoch.startDate)
  if (Number.isNaN(d0.getTime())) return null
  return Array.from({ length: MONTH_COUNT }, (_, i) => {
    const d = new Date(d0)
    d.setUTCMonth(d.getUTCMonth() + i)
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`
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

function splitCsvLine(line) {
  return line.split(',').map((c) => c.trim())
}

function toRecord(cols) {
  return Object.fromEntries(HEADERS.map((h, i) => [h, cols[i] ?? '']))
}

function describeKey(r) {
  return `${r.country_iso3} / ${r.month_date} / ${r.indicator} / ${r.segment}`
}

export function fileError(message) {
  return { row: null, column: null, message, record: null, text: message }
}

export function emptyCsvResult(message) {
  return { ok: false, errors: [fileError(message)], rows: [], preview: [], missingCells: [], dataRowCount: 0 }
}

export function rewriteSampleMonths(text, months) {
  const lines = String(text).replace(/^\uFEFF/, '').split(/\r?\n/)
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

export function parseAndValidateCsv(text, months) {
  const errors = []
  const add = (row, column, message, record = null, extra = {}) => {
    const where = row == null ? '' : column ? `Row ${row}, column ${column}: ` : `Row ${row}: `
    errors.push({ row, column, message, record, ...extra, text: `${where}${message}` })
  }

  if (!months || months.length !== MONTH_COUNT) {
    return emptyCsvResult('Epoch months are not available; cannot validate CSV dates')
  }

  if (text == null || String(text).trim() === '') {
    return emptyCsvResult('File is empty')
  }

  const rawLines = String(text).replace(/^\uFEFF/, '').split(/\r?\n/)
  while (rawLines.length > 0 && rawLines[rawLines.length - 1].trim() === '') {
    rawLines.pop()
  }

  if (rawLines.length === 0) {
    return emptyCsvResult('File is empty')
  }

  const header = splitCsvLine(rawLines[0])
  if (header.length !== HEADERS.length || header.some((h, i) => h !== HEADERS[i])) {
    add(1, null, `Header must be exactly ${HEADERS.join(',')} (got ${header.join(',')})`)
  }
  if (rawLines.length < 2) {
    add(null, null, 'File has a header but no data rows')
  }

  const monthSet = new Set(months)
  const parsed = []
  const seen = new Map()
  const foreignMonths = new Set()
  let dataRowCount = 0
  let duplicateCount = 0

  for (let i = 1; i < rawLines.length; i++) {
    const rowNum = i + 1
    const line = rawLines[i]
    if (line.trim() === '') {
      add(rowNum, null, 'Blank line is not allowed')
      continue
    }
    dataRowCount++
    const cols = splitCsvLine(line)
    const record = toRecord(cols)

    if (cols.length < HEADERS.length) {
      const absent = HEADERS.slice(cols.length)
      add(rowNum, absent[0], `Expected ${HEADERS.length} columns, found ${cols.length} (missing ${absent.join(', ')})`, record)
      continue
    }
    if (cols.length > HEADERS.length) {
      add(rowNum, null, `Expected ${HEADERS.length} columns, found ${cols.length}`, record)
      continue
    }

    let rowOk = true
    const fail = (column, message) => {
      rowOk = false
      add(rowNum, column, message, record)
    }

    const empty = HEADERS.filter((h) => record[h] === '')
    for (const h of empty) fail(h, `${h} is empty`)

    const { month_date: monthDate, country_iso3: country, indicator, segment, value: valueRaw } = record

    if (monthDate !== '') {
      if (!MONTH_RE.test(monthDate)) {
        fail('month_date', `month_date must be YYYY-MM (got "${monthDate}")`)
      } else if (!monthSet.has(monthDate)) {
        foreignMonths.add(monthDate)
        fail('month_date', `"${monthDate}" is not an epoch month; expected one of ${months.join(', ')}`)
      }
    }
    if (country !== '' && !COUNTRY_SET.has(country)) {
      fail('country_iso3', `country_iso3 must be one of ${COUNTRIES.join(', ')} (got "${country}")`)
    }
    if (indicator !== '' && segment !== '' && !SERIES_SET.has(`${indicator}|${segment}`)) {
      fail(INDICATOR_SET.has(indicator) ? 'segment' : 'indicator', `Unknown indicator/segment pair "${indicator}","${segment}"`)
    }

    let value = null
    if (valueRaw !== '') {
      if (!VALUE_RE.test(valueRaw)) {
        fail('value', `value must be a positive integer with no sign, decimal, or separators (got "${valueRaw}")`)
      } else {
        value = BigInt(valueRaw)
        if (value === 0n) fail('value', 'value cannot be 0')
        else if (value > LONG_MAX) fail('value', 'value exceeds the maximum 64-bit integer')
      }
    }

    if (!rowOk) continue

    const key = `${country}|${monthDate}|${indicator}|${segment}`
    const original = seen.get(key)
    if (original) {
      duplicateCount++
      add(rowNum, null, `Duplicate of row ${original.row}: ${describeKey(record)}`, record, { duplicateOf: original })
      continue
    }
    seen.set(key, { row: rowNum, record })

    parsed.push({ rowNum, month: monthDate, country, indicator, segment, value })
  }

  const presentMonths = new Set(parsed.map((r) => r.month))
  const missingMonths = months.filter((m) => !presentMonths.has(m))
  if (foreignMonths.size > 0) {
    add(null, null, `month_date values ${[...foreignMonths].sort().join(', ')} do not belong to this epoch; use ${months.join(', ')}`)
  }
  if (dataRowCount > 0 && missingMonths.length > 0) {
    add(null, null, `Missing month(s): ${missingMonths.join(', ')}`)
  }

  const missingMonthSet = new Set(missingMonths)
  const missingCells = []
  for (const country of COUNTRIES) {
    for (const month of months) {
      if (missingMonthSet.has(month)) continue
      for (const [indicator, segment] of SERIES) {
        if (!seen.has(`${country}|${month}|${indicator}|${segment}`)) {
          missingCells.push({ country, month, indicator, segment })
        }
      }
    }
  }

  if (dataRowCount > 0 && (dataRowCount !== CELL_COUNT || duplicateCount > 0 || missingCells.length > 0 || missingMonths.length > 0)) {
    const parts = [`File has ${dataRowCount} data row(s) (expected ${CELL_COUNT}): ${parsed.length} valid unique`]
    if (duplicateCount > 0) parts.push(`${duplicateCount} duplicate`)
    const invalid = dataRowCount - parsed.length - duplicateCount
    if (invalid > 0) parts.push(`${invalid} invalid`)
    let summary = parts.join(', ') + '.'
    const missingTotal = CELL_COUNT - parsed.length
    if (missingTotal > 0) summary += ` ${missingTotal} required cell(s) not covered.`
    add(null, null, summary)
  }

  const rows = parsed

  const preview = rows.slice(0, 5).map((r) => ({
    country: r.country,
    month: r.month,
    indicator: r.indicator,
    segment: r.segment,
    value: r.value.toString(),
  }))

  return {
    ok: errors.length === 0 && rows.length === CELL_COUNT,
    errors,
    rows,
    preview,
    missingCells,
    dataRowCount,
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
