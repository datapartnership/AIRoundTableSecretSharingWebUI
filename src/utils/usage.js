// Submission streams; keys match the API's UsageTypes.
export const USAGE_TYPES = ['api', 'chat']
export const USAGE = {
  api: { label: 'API usage', nav: 'API Usage' },
  chat: { label: 'Consumer Chat Usage', nav: 'Consumer Chat Usage' },
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

// Epoch and period dates are first-of-month UTC
function ym(iso) {
  const m = /^(\d{4})-(\d{2})/.exec(String(iso ?? ''))
  return m ? [Number(m[1]), Number(m[2]) - 1] : null
}

/** "August 2026" */
export function monthYear(iso, addMonths = 0) {
  const p = ym(iso)
  if (!p) return ''
  const t = p[0] * 12 + p[1] + addMonths
  return `${MONTHS[t % 12]} ${Math.floor(t / 12)}`
}

/** "August to November 2026", or "November 2026 to February 2027" across a year end */
export function monthRange(iso, count) {
  const p = ym(iso)
  if (!p) return ''
  const end = p[0] * 12 + p[1] + count - 1
  const endYear = Math.floor(end / 12)
  return endYear === p[0]
    ? `${MONTHS[p[1]]} to ${MONTHS[end % 12]} ${endYear}`
    : `${monthYear(iso)} to ${MONTHS[end % 12]} ${endYear}`
}

/** "22 September 2026" */
export function longDate(iso) {
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`
}

/** "14:32" */
export function clockTime(d = new Date()) {
  return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`
}

export const SUPPORT_EMAIL = 'datapartnership@worldbank.org'
