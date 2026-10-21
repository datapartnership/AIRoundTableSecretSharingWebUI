// Turns validator output (csvUpload.js) into the file-check screens of the wireframe (page 06, S1 to S18):
// a summary line, problem groups with what / where / accepted / fix and example rows, or a Blocked state.
import { COUNTRIES, HEADERS, INDICATORS, SEGMENTS_BY_INDICATOR, SERIES, CELL_COUNT } from './csvUpload'
import { monthRange } from './usage'

export const EXAMPLE_COLS = ['row', ...HEADERS]
const MISSING_COLS = ['month_date', 'country_iso3', 'indicator', 'segment']
const MISSING_SUMMARY_COLS = ['country_iso3', 'month_date', 'series', 'missing rows']
const HEADER = HEADERS.join(',')

const ACC_NUM = 'Accepted: numeric values above 0 without special characters.'
const ACC_ALL = `Accepted: one row for each country, month, and series: ${COUNTRIES.length} countries x 4 months x ${SERIES.length} series = ${CELL_COUNT} rows.`
const ACC_DUP = 'Accepted: each combination of month, country, indicator, and segment one time only.'
const FIX_DUP = 'Keep one row for this combination, with the correct value, and delete the other.'
const FIX_MISS = 'Add one row for each missing combination, with its value.'

const plural = (n, one, many) => (n === 1 ? one : many)
const listAnd = (items) => items.length <= 1 ? items.join('') : items.length === 2 ? `${items[0]} and ${items[1]}` : `${items.slice(0, -1).join(', ')}, and ${items[items.length - 1]}`

// "rows 4, 9, 22, and 11 more" / "row 57" / "all 200 rows"
function rowList(rows, total) {
  if (rows.length === total && total > 3) return `all ${total} rows`
  if (rows.length === 1) return `row ${rows[0]}`
  if (rows.length <= 3) return `rows ${listAnd(rows.map(String))}`
  return `rows ${rows.slice(0, 3).join(', ')}, and ${rows.length - 3} more`
}

// "6.15498E+15, 8.42531E+15, and 9 more" for values; "open-source (20)" for names
function foundValues(found, withCounts = false) {
  const show = (v) => (v === '' ? 'an empty value' : v)
  if (withCounts) {
    const counts = new Map()
    for (const f of found) counts.set(f, (counts.get(f) ?? 0) + 1)
    return [...counts].map(([v, n]) => (n > 1 ? `${show(v)} (${n})` : show(v))).join(', ')
  }
  const distinct = [...new Set(found)]
  if (distinct.length <= 5) return distinct.map(show).join(', ')
  return `${distinct.slice(0, 5).map(show).join(', ')}, and ${distinct.length - 5} more`
}

const exampleRow = (e) => [String(e.row), ...HEADERS.map((h) => (e.record?.[h] === '' ? '(empty)' : e.record?.[h] ?? ''))]
const badIndex = (column) => (column ? EXAMPLE_COLS.indexOf(column) : -1)

function group(what, where, accepted, fix, errors, column, extra = {}) {
  return {
    what, where, accepted, fix,
    cols: EXAMPLE_COLS,
    rows: errors.map(exampleRow),
    bad: badIndex(column),
    report: errors.map((e) => [e.row, column ?? '', e.kind === 'columns' ? `${e.found} values` : e.found]),
    ...extra,
  }
}

// Missing combinations: listed one by one when few, otherwise summarised per country and month
function missingGroup(missing, months, found) {
  const n = missing.length
  const what = `${n} ${plural(n, 'combination', 'combinations')}${n <= 3 ? ' of country, month, and series' : ''} ${plural(n, 'is', 'are')} missing`
  if (n <= 3) {
    return {
      what,
      report: missing.map((c) => ['', '', `${c.month} ${c.country} ${c.indicator} ${c.segment}`]),
      where: `Found: ${found} of ${CELL_COUNT} combinations. The ${n === 1 ? 'missing combination is' : `${n} missing combinations are`} below.`,
      accepted: ACC_ALL, fix: FIX_MISS, cols: MISSING_COLS, bad: -1,
      rows: missing.map((c) => [c.month, c.country, c.indicator, c.segment]),
    }
  }
  const byCountry = new Map()
  for (const c of missing) {
    const perMonth = byCountry.get(c.country) ?? new Map()
    perMonth.set(c.month, (perMonth.get(c.month) ?? 0) + 1)
    byCountry.set(c.country, perMonth)
  }
  const rows = []
  for (const [country, perMonth] of byCountry) {
    const full = [...perMonth.values()].every((k) => k === SERIES.length)
    if (full && perMonth.size === months.length) {
      rows.push([country, `${months[0]} to ${months[months.length - 1]}`, `all ${SERIES.length} series`, String(perMonth.size * SERIES.length)])
    } else {
      for (const [month, k] of perMonth) rows.push([country, month, k === SERIES.length ? `all ${SERIES.length} series` : `${k} of ${SERIES.length} series`, String(k)])
    }
  }
  const columns = [...new Set([byCountry.size < COUNTRIES.length && 'country_iso3', rows.some((r) => !r[1].includes(' to ')) && 'month_date'].filter(Boolean))]
  return {
    what,
    report: missing.map((c) => ['', '', `${c.month} ${c.country} ${c.indicator} ${c.segment}`]),
    where: `${columns.length ? `Column${columns.length > 1 ? 's' : ''} ${listAnd(columns)}. ` : ''}Found: ${found} of ${CELL_COUNT} combinations. The ${n} missing rows are summarized below. The problem report lists each row.`,
    accepted: ACC_ALL, fix: FIX_MISS, cols: MISSING_SUMMARY_COLS, bad: -1, rows,
  }
}

/**
 * Builds the file-check view for a validation result. Returns
 * { blocked, big, sub, groups, text, acts, compact, meta }.
 */
export function buildFileCheck(result, months, fileName) {
  const fe = result.fileError
  if (fe) return blockedView(fe, fileName)

  const { errors, missingCells, dataRowCount } = result
  const by = (kind) => errors.filter((e) => e.kind === kind)
  const groups = []
  const where = (column, list) => `Column ${column}, ${rowList(list.map((e) => e.row), dataRowCount)}`
  const epochMonthsText = `Accepted: ${months.join(', ')}.`

  const empty = by('value-empty')
  if (empty.length) {
    groups.push(group(`${empty.length} ${plural(empty.length, 'value is', 'values are')} empty`,
      `${where('value', empty)}. Found: ${plural(empty.length, 'an empty value', 'empty values')}.`, ACC_NUM, 'Enter the value for each row.', empty, 'value'))
  }

  const bad = by('value')
  if (bad.length) {
    // A ticked metric with 0 needs checking at the source; other values need rewriting
    const zeros = bad.every((e) => /^-?[0-9]*\.?[0-9]*$/.test(e.found))
    groups.push(group(`${bad.length} ${plural(bad.length, 'value is not a whole number', 'values are not whole numbers')} above 0`,
      `${where('value', bad)}. Found: ${foundValues(bad.map((e) => e.found))}.`,
      bad.some((e) => e.rule === 'positive') ? `${ACC_NUM} You ticked these metrics in the Quorum Check, so their values must be above 0.` : ACC_NUM,
      zeros ? 'Check these values in your source system. If a real value is 0, ask the platform administrator how to report it.' : 'Write each value with digits only.',
      bad, 'value'))
  }

  const notZero = by('value-not-zero')
  if (notZero.length) {
    groups.push(group(`${notZero.length} ${plural(notZero.length, 'value is', 'values are')} not 0 for a metric you did not tick`,
      `${where('value', notZero)}. Found: ${foundValues(notZero.map((e) => e.found))}.`,
      'Accepted: 0 for each metric you did not tick in the Quorum Check.',
      'Enter 0 for these rows. If your Quorum Check answers were wrong, contact the platform administrator.', notZero, 'value'))
  }

  const fmt = by('month-format')
  if (fmt.length) {
    groups.push(group(`${fmt.length} month ${plural(fmt.length, 'value is', 'values are')} not written as YYYY-MM`,
      `${where('month_date', fmt)}. Found: ${foundValues(fmt.map((e) => e.found))}.`,
      `Accepted: YYYY-MM. For this epoch: ${months.join(', ')}.`, 'Write each month as YYYY-MM.', fmt, 'month_date'))
  }

  const foreign = by('month-foreign')
  if (foreign.length) {
    groups.push(group(`${foreign.length} ${plural(foreign.length, 'row uses a month', 'rows use months')} outside this epoch`,
      `${where('month_date', foreign)}. Found: ${foundValues(foreign.map((e) => e.found))}.`,
      epochMonthsText, `Use the values for ${monthRange(`${months[0]}-01`, months.length)}, with these ${months.length} months.`, foreign, 'month_date'))
  }

  const country = by('country')
  if (country.length) {
    groups.push(group(`${country.length} ${plural(country.length, 'row has a country code', 'rows have a country code')} that is not accepted`,
      `${where('country_iso3', country)}. Found: ${foundValues(country.map((e) => e.found), true)}.`,
      `Accepted: country codes in ISO 3166-1 alpha-3 only: ${COUNTRIES.join(', ')}.`, 'Use one of the accepted codes.', country, 'country_iso3'))
  }

  const segment = by('segment')
  if (segment.length) {
    const inds = [...new Set(segment.map((e) => e.record.indicator))]
    groups.push(group(`${segment.length} ${plural(segment.length, 'row has a segment', 'rows have a segment')} that is not accepted`,
      `${where('segment', segment)}. Found: ${foundValues(segment.map((e) => e.found), true)}.`,
      inds.map((i) => `Accepted for ${i}: ${SEGMENTS_BY_INDICATOR[i].join(', ')}.`).join(' '), 'Use one of the accepted segments.', segment, 'segment'))
  }

  const indicator = by('indicator')
  if (indicator.length) {
    groups.push(group(`${indicator.length} ${plural(indicator.length, 'row has an indicator', 'rows have an indicator')} that is not accepted`,
      `${where('indicator', indicator)}. Found: ${foundValues(indicator.map((e) => e.found), true)}.`,
      `Accepted: ${INDICATORS.join(', ')}. Names are case-sensitive.`, 'Use one of the accepted indicators.', indicator, 'indicator'))
  }

  for (const [kind, same] of [['duplicate-same', true], ['duplicate-diff', false]]) {
    const dups = by(kind)
    if (!dups.length) continue
    const pairs = dups.flatMap((e) => [{ row: e.duplicateOf.row, record: e.duplicateOf.record }, e])
    groups.push(group(`${dups.length} ${plural(dups.length, 'combination appears', 'combinations appear')} twice, with ${same ? 'the same values' : 'different values'}`,
      `Columns month_date, country_iso3, indicator, and segment, ${rowList(pairs.map((e) => e.row), -1)}. Found: the same combination two times, with ${same ? 'the same value' : 'different values'}.`,
      ACC_DUP, FIX_DUP, pairs, same ? null : 'value'))
  }

  const cols = by('columns')
  if (cols.length) {
    const rows = cols.map((e) => e.row)
    groups.push(group(`${cols.length} ${plural(cols.length, 'row does', 'rows do')} not have ${HEADERS.length} values`,
      `${rowList(rows, dataRowCount).replace(/^r/, 'R')}. Found: ${listAnd(cols.slice(0, 3).map((e) => `${e.found} values in row ${e.row}`))}${cols.length > 3 ? `, and ${cols.length - 3} more` : ''}.`,
      `Accepted: ${HEADERS.length} values in each row, one for each column in the header, with one comma between values.`,
      `Make each row match the header: ${HEADERS.length} values, with one comma between them.`,
      cols.map((e) => ({ ...e, record: Object.fromEntries(HEADERS.map((h, i) => [h, e.record[h] || (i >= e.found ? '(none)' : '')])) })), null))
  }

  if (missingCells.length) groups.push(missingGroup(missingCells, months, CELL_COUNT - missingCells.length))

  // The summary line, as on page 06
  const parts = []
  const badN = result.badRowCount
  if (badN > 0) parts.push(badN === dataRowCount && badN > 1 ? `All ${badN} rows need correction.` : `${badN} ${plural(badN, 'row needs', 'rows need')} correction.`)
  if (missingCells.length) parts.push(`${missingCells.length} ${plural(missingCells.length, 'row is', 'rows are')} missing.`)
  if (result.readyCount > 0 && badN < dataRowCount) parts.push(`${result.readyCount} ${plural(result.readyCount, 'row is', 'rows are')} ready.`)
  const onlyMissing = badN === 0
  const sub = onlyMissing
    ? `Add the missing ${plural(missingCells.length, 'row', 'rows')}, then choose the file again.`
    : groups.length === 1 ? 'Fix the problem below in your file, then upload it again.' : `Fix the ${groups.length} problems below in your file, then upload it again.`

  return {
    blocked: false,
    big: parts.join(' '),
    sub,
    groups,
    // More than 3 groups: each group shows only what, where, and the fix; example rows open on request
    compact: groups.length > 3,
    meta: `${dataRowCount} rows | Checked in this browser at `,
  }
}

function blockedView(fe, fileName) {
  switch (fe.kind) {
    case 'not-csv':
      return {
        blocked: true, meta: 'Not checked', big: 'This file can’t be checked.',
        groups: [{ what: 'The file type is not accepted', where: `File type. Found: ${fileExt(fileName)}.`, accepted: 'Accepted: .csv files only: plain text, with one comma between values.', fix: 'Save the file as CSV UTF-8 (Comma delimited). Then choose the .csv file here.', rows: [] }],
        acts: [['choose', 'Choose a CSV file'], ['sample', 'Download sample.csv']],
      }
    case 'header':
      return {
        blocked: true, meta: `${fe.rows} rows, including the header`, big: 'The rows were not checked, because the columns don’t match.',
        groups: [{ what: 'The header is not accepted', where: 'Row 1, the header. Each column is found by its exact name, so the header must match exactly.', accepted: '',
          code: [['Found:', fe.found], ['Accepted:', HEADER]], copyHeader: true,
          fix: 'Replace row 1 with the accepted header, and delete any other columns. Or start again from the sample file.', rows: [] }],
        acts: [['choose', 'Upload corrected file'], ['sample', 'Download sample.csv']],
      }
    case 'separator':
      return {
        blocked: true, meta: 'Not checked', big: 'The rows were not checked, because the values are not separated by commas.',
        groups: [{ what: 'The separator is not accepted', where: 'All rows. The file must have a comma between values.', accepted: '',
          code: [['Found:', fe.found], ['Accepted:', HEADER]],
          fix: 'Export the file again with a comma between values. In Excel, the separator follows the list separator in your computer’s region settings.', rows: [] }],
        acts: [['choose', 'Upload corrected file'], ['sample', 'Download sample.csv']],
      }
    case 'unreadable':
      return {
        blocked: true, meta: 'Not checked', big: 'This browser couldn’t open the file.',
        text: ['The browser can open only a file that is saved on this computer and not open in another program.',
          'Close the file in any other program. Make sure it is saved on this computer, not only in a cloud folder. Then choose it again.'],
        acts: [['choose', 'Choose the file again']],
      }
    case 'empty':
      return {
        blocked: true, meta: 'Not checked', big: 'This file is empty.',
        text: [`The file must start with the header ${HEADER}, then one row for each country, month, and series.`],
        acts: [['choose', 'Choose another file'], ['sample', 'Download sample.csv']],
      }
    default:
      return {
        blocked: true, meta: 'Not checked', big: 'The platform couldn’t check your file.',
        text: ['The problem is on the platform, not in your file. Do not change your file.',
          'Reload the page, then choose the file again. If this message comes back, contact the platform administrator.'],
        acts: [['reload', 'Reload the page']],
      }
  }
}

function fileExt(name) {
  const m = /\.([^.]+)$/.exec(name ?? '')
  if (!m) return 'a file without an extension'
  const ext = m[1].toLowerCase()
  return ext === 'xlsx' || ext === 'xls' ? `.${ext}, an Excel workbook` : `.${ext}`
}

const csvq = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`

/** The downloadable problem report: one line for each row problem and each missing combination. */
export function problemReportCsv(view) {
  const lines = ['row,column,found,problem,accepted,fix']
  for (const g of view.groups ?? []) {
    for (const [row, column, found] of g.report ?? [['', '', '']]) {
      lines.push([row, column, found, g.what, g.accepted, g.fix].map(csvq).join(','))
    }
  }
  return lines.join('\n') + '\n'
}
