import { useEffect, useRef, useState } from 'react'
import { Link, useSearchParams } from 'react-router-dom'
import { Page, announce, useShell } from '../components/Shell'
import { Alert, epochTag } from '../components/ui'
import useEpochs, { byStartDesc, contribution } from '../hooks/useEpochs'
import { USAGE, USAGE_TYPES, monthYear } from '../utils/usage'

// The same words as the Home rows; also the options of the Your contribution filter
const MINE = { todo: 'Please share your data', done: 'Your masked data were submitted' }
const FILTERS = [['year', 'Year', 'All years'], ['u', 'Usage type', 'All usage types'], ['status', 'Status', 'All statuses'], ['mine', 'Your contribution', 'All']]
const statusOf = (e) => (e.cancelledAt ? 'cancelled' : e.isClosed ? 'closed' : 'open')

const toRow = (e) => ({
  e,
  month: monthYear(e.startDate),
  year: String(e.startDate).slice(0, 4),
  u: e.usageType,
  status: statusOf(e),
  mine: contribution(e),
})

function optionsFor(k, rows) {
  if (k === 'year') return [...new Set(rows.map((r) => r.year))].map((y) => [y, y])
  if (k === 'u') return USAGE_TYPES.map((u) => [u, USAGE[u].label])
  if (k === 'status') {
    const base = [['open', 'Open'], ['closed', 'Closed']]
    return rows.some((r) => r.status === 'cancelled') ? [...base, ['cancelled', 'Cancelled']] : base
  }
  return Object.entries(MINE)
}

export default function AllEpochs() {
  const { isAdmin } = useShell()
  const { epochs, error, loading, retry } = useEpochs({ poll: false })
  const [params, setParams] = useSearchParams()
  const [showFilters, setShowFilters] = useState(false)
  const h1 = useRef(null)

  const all = (epochs ?? []).filter((e) => isAdmin || e.isEligible).sort(byStartDesc).map(toRow)
  // Accept only known filters and values from the page address
  const f = {}
  for (const [k] of FILTERS) {
    const v = params.get(k)
    if (v && optionsFor(k, all).some(([o]) => o === v)) f[k] = v
  }
  const rows = all.filter((r) => FILTERS.every(([k]) => !f[k] || r[k] === f[k]))
  const set = FILTERS.filter(([k]) => f[k]).map(([k]) => optionsFor(k, all).find(([v]) => v === f[k])[1])
  const n = Object.keys(f).length
  const countText = `Showing ${rows.length} of ${all.length} epochs${set.length ? ': ' + set.join(', ') : ''}`

  // Filters apply at once; the select keeps the focus and the count is announced
  const lastCount = useRef(null)
  useEffect(() => {
    if (loading || error) return
    if (lastCount.current !== null && lastCount.current !== countText) announce(countText)
    lastCount.current = countText
  }, [countText, loading, error])

  const setFilter = (k, v) => {
    const next = new URLSearchParams(params)
    if (v) next.set(k, v); else next.delete(k)
    setParams(next, { replace: true })
  }
  const clear = () => {
    setParams({}, { replace: true })
    h1.current?.focus()
  }

  const btn = (r) => !isAdmin && r.status === 'open' && r.mine === 'todo'
    ? <Link className="btn btn-primary" to={`/${r.u}?epoch=${r.e.epochId}`} aria-label={`Continue ${USAGE[r.u].label}, ${r.month}`}>Continue</Link>
    : null
  const note = (r) => (r.status === 'cancelled' ? 'Cancelled by an admin' : isAdmin ? `${r.e.producerCount} partners` : MINE[r.mine])
  const empty = `No epochs match these filters. Change a filter, or clear the filters to see all ${all.length} epochs.`

  let body
  if (loading) {
    body = <div role="status"><p className="ep-status">Checking your epochs</p><p className="meta" style={{ color: 'var(--text-secondary)' }}>This takes a moment.</p></div>
  } else if (error) {
    body = (
      <div className="ep-filters" style={{ gap: 16 }}>
        <Alert kind="error" title="Your epochs could not be loaded" lines={['Your work is not affected. Try again, or contact the platform administrator if this continues.']} />
        <button className="btn btn-primary" style={{ width: 'auto' }} onClick={retry}>Try again</button>
      </div>
    )
  } else {
    body = (
      <div className="ep-filters">
        <button className="btn btn-secondary ep-toggle" aria-expanded={showFilters} aria-controls="ep-fields" onClick={() => setShowFilters(!showFilters)}>
          {showFilters ? 'Hide filters' : 'Show filters'}{n ? ` (${n})` : ''}
        </button>
        <div className={`ep-fields${showFilters ? ' open' : ''}`} id="ep-fields">
          {FILTERS.map(([k, label, any]) => (
            <label key={k} className="field">{label}
              <select name={k} value={f[k] ?? ''} onChange={(ev) => setFilter(k, ev.target.value)}>
                <option value="">{any}</option>
                {optionsFor(k, all).map(([v, t]) => <option key={v} value={v}>{t}</option>)}
              </select>
            </label>
          ))}
          <span className="sp" />
        </div>
        <div>
          <p className="ep-count">
            <span>{countText}</span>
            {set.length > 0 && <button className="link" onClick={clear}>Clear filters</button>}
          </p>
          <table className="ep-table">
            <colgroup><col /><col style={{ width: 212 }} /><col style={{ width: 144 }} /><col style={{ width: 292 }} /><col style={{ width: 118 }} /></colgroup>
            <thead><tr><th scope="col">Epoch</th><th scope="col">Usage type</th><th scope="col">Status</th><th scope="col">{isAdmin ? 'Partners' : 'Your contribution'}</th><th scope="col"><span className="sr-only">Action</span></th></tr></thead>
            <tbody>
              {rows.length ? rows.map((r) => (
                <tr key={r.e.epochId}><td className="month">{r.month}</td><td>{USAGE[r.u].label}</td><td>{epochTag(r.e)}</td><td>{note(r)}</td><td>{btn(r)}</td></tr>
              )) : <tr><td colSpan={5} className="ep-empty">{empty}</td></tr>}
            </tbody>
          </table>
          <ul className="ep-rows">
            {rows.length ? rows.map((r) => (
              <li key={r.e.epochId} className="erow">
                <div className="info"><div className="name"><span className="month">{r.month}</span>{epochTag(r.e)}</div><p className="rnote">{USAGE[r.u].label}<br />{note(r)}</p></div>
                {btn(r)}
              </li>
            )) : <li className="erow"><p>{empty}</p></li>}
          </ul>
        </div>
      </div>
    )
  }

  return (
    <Page>
      <main id="main" className="allep wrap">
        <Link className="link back" to="/">← Home</Link>
        <div className="epoch-head"><h1 ref={h1} tabIndex={-1}>All epochs</h1>{!isAdmin && <OrgLine />}</div>
        {body}
      </main>
    </Page>
  )
}

function OrgLine() {
  const { user } = useShell()
  return user?.org ? <p className="meta">Your organization: {user.org}</p> : null
}
