import { useCallback, useEffect, useRef, useState } from 'react'
import { Link, useSearchParams } from 'react-router-dom'
import { useMsal } from '@azure/msal-react'
import { Page, announce, download, useShell } from '../components/Shell'
import { Alert, ProtocolStrip, Tag, epochTag, stepAnnouncement } from '../components/ui'
import useEpochs, { byStartDesc } from '../hooks/useEpochs'
import useKeyExchange from '../hooks/useKeyExchange'
import * as api from '../utils/api'
import { calculateMaskedValue } from '../utils/noise'
import { getDeviceId } from '../utils/localCrypto'
import {
  CELL_COUNT, COUNTRIES, HEADERS, SEGMENTS_BY_INDICATOR, cellKey, epochMonths, isCsvFile, pairKey,
  parseAndValidateCsv, rewriteSampleMonths, valueRule,
} from '../utils/csvUpload'
import { buildFileCheck, problemReportCsv } from '../utils/fileCheck'
import { answersMap, loadQuorumLocal } from '../utils/quorumLocal'
import { USAGE, SUPPORT_EMAIL, clockTime, longDate, monthRange, monthYear } from '../utils/usage'

const POLL_MS = 5000
const D = 'done', C = 'current', W = 'waiting', P = 'problem', L = 'later'
const SEC_FILE = 'Upload and verify data'
const SEC_SUBMIT = 'Masked data submission'
const SUBMIT_P = 'Your browser masks each value, then sends only the masked values. Submitting is final: you cannot change your contribution after this.'
const ACCESS = SEGMENTS_BY_INDICATOR.token_access
const REGIONS = SEGMENTS_BY_INDICATOR.model_region_of_origin
const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`

export default function EpochPage({ usageType: u }) {
  const { isAdmin } = useShell()
  const { epochs, error, loading, retry } = useEpochs()
  const [params] = useSearchParams()
  const requested = Number(params.get('epoch')) || null

  const mine = (epochs ?? []).filter((e) => e.usageType === u && (isAdmin || e.isEligible)).sort(byStartDesc)
  const epoch = mine.find((e) => e.epochId === requested)
    ?? mine.find((e) => !e.isClosed && !e.cancelledAt)
    ?? mine.find((e) => !e.cancelledAt)
    ?? null

  let body
  if (loading) body = <EpochShell u={u}><div role="status"><p className="meta">Checking your epochs. This takes a moment.</p></div></EpochShell>
  else if (error) {
    body = (
      <EpochShell u={u}>
        <Alert kind="error" title="Your epochs could not be loaded" lines={['Your work is not affected. Try again, or contact the platform administrator if this continues.']}>
          <div className="actions"><button className="btn btn-secondary" onClick={retry}>Try again</button></div>
        </Alert>
      </EpochShell>
    )
  } else if (!epoch) {
    // Not drawn: proposal for a usage type without any epoch yet
    body = <EpochShell u={u}><p className="page-note">No {USAGE[u].label} epoch is open right now. The next epoch will appear on <Link className="link" to="/">Home</Link> when it opens.</p></EpochShell>
  } else body = <EpochView key={epoch.epochId} u={u} epoch={epoch} isAdmin={isAdmin} />

  return <Page>{body}</Page>
}

function EpochShell({ u, epoch, children }) {
  return (
    <main id="main" className="epoch wrap">
      <div className="epoch-head">
        <p className="usage">{USAGE[u].label}</p>
        <div className="title-row">
          <h1 tabIndex={-1}>{epoch ? `${monthYear(epoch.startDate)} epoch` : USAGE[u].label}</h1>
          {epoch && epochTag(epoch)}
        </div>
        {epoch && <p className="meta">Reporting months: {monthRange(epoch.startDate, 4)}</p>}
      </div>
      {children}
    </main>
  )
}

const Sec = ({ h, children }) => <section className="sec"><h2 tabIndex={-1}>{h}</h2>{children}</section>
const PartNow = ({ children }) => <Sec h="Your part now">{children}</Sec>

function EpochView({ u, epoch, isAdmin }) {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  const myId = account?.localAccountId ?? ''
  const deviceId = getDeviceId()
  const epochId = epoch.epochId
  const months = epochMonths(epoch) ?? []
  const viewOnly = isAdmin || !epoch.isEligible
  const inactive = viewOnly || epoch.isClosed || !!epoch.cancelledAt

  // A finished contribution needs no key exchange
  const kx = useKeyExchange({ scope: { epochId }, producerIds: epoch.producerIds ?? [], active: !inactive && !epoch.mySubmissionComplete })

  // ── Quorum Check period of this epoch ────────────────────────────────────────
  const [quorum, setQuorum] = useState(epoch.quorumPeriodId ? null : { quorumComplete: true, ignoredMetrics: [] })
  useEffect(() => {
    if (!epoch.quorumPeriodId || inactive) return
    let alive = true
    const load = async () => {
      try {
        const token = await api.acquireApiToken(instance, account)
        const q = await api.getQuorumPeriod(epoch.quorumPeriodId, token)
        if (alive) setQuorum(q)
      } catch (e) {
        console.error('[epoch] quorum status failed', e)
      }
    }
    load()
    if (quorum?.quorumComplete) return () => { alive = false }
    const id = setInterval(load, POLL_MS)
    return () => { alive = false; clearInterval(id) }
  }, [epoch.quorumPeriodId, inactive, quorum?.quorumComplete, instance, account?.homeAccountId])

  const quorumRules = {
    answers: epoch.quorumPeriodId ? answersMap(loadQuorumLocal(myId, epoch.quorumPeriodId, deviceId)) : null,
    ignoredPairs: new Set((quorum?.ignoredMetrics ?? []).map((m) => pairKey(m.country, m.indicator))),
  }

  // ── The partner's stored submissions ─────────────────────────────────────────
  const [submitted, setSubmitted] = useState(null)   // { keys: Set, at: Date | null }
  const loadSubmissions = useCallback(async () => {
    const token = await api.acquireApiToken(instance, account)
    const my = await api.getMySubmissions(epochId, token)
    const list = my.submissions ?? []
    const at = list.map((s) => s.submittedAt).filter(Boolean).sort().pop()
    const value = { keys: new Set(list.map((s) => cellKey(s.country, s.month, s.indicator, s.segment))), at: at ? new Date(at) : null }
    setSubmitted(value)
    return value
  }, [instance, account?.homeAccountId, epochId])

  useEffect(() => {
    if (viewOnly || epoch.cancelledAt || !(kx.ready || epoch.isClosed || epoch.mySubmissionComplete)) return
    loadSubmissions().catch((e) => console.error('[epoch] load submissions failed', e))
  }, [kx.ready, epoch.isClosed, epoch.mySubmissionComplete, viewOnly, loadSubmissions])

  // Metrics below the quorum get no data, so the partner owes the rest
  const ignoredCellCount = months.length * [...quorumRules.ignoredPairs].reduce((n, k) => n + SEGMENTS_BY_INDICATOR[k.split('|')[1]].length, 0)
  const owed = CELL_COUNT - ignoredCellCount
  const belowQuorumText = [...quorumRules.ignoredPairs].map((k) => k.replace('|', ' ')).join(', ')
  const allSubmitted = epoch.mySubmissionComplete || (submitted && submitted.keys.size >= owed && owed > 0)

  // ── Upload and submit ────────────────────────────────────────────────────────
  const [upload, setUpload] = useState({ phase: 'idle' })   // idle | checking | checked
  const [submit, setSubmit] = useState({ phase: 'idle' })   // idle | submitting | failed
  const fileInput = useRef(null)

  const choose = () => fileInput.current?.click()

  const onFile = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    setSubmit({ phase: 'idle' })
    setUpload({ phase: 'checking', fileName: file.name })
    const at = new Date()
    let result
    if (!isCsvFile(file)) result = { fileError: { kind: 'not-csv' } }
    else {
      let text = null
      try { text = await file.text() } catch { result = { fileError: { kind: 'unreadable' } } }
      if (text !== null) {
        try { result = parseAndValidateCsv(text, months, quorumRules) } catch (e) {
          console.error('[epoch] csv check failed', e)
          result = { fileError: { kind: 'platform' } }
        }
      }
    }
    // Let the "Checking" state render before the result replaces it
    setTimeout(() => setUpload({ phase: 'checked', fileName: file.name, at, result, view: buildFileCheck(result, months, file.name) }), 400)
  }

  const downloadSample = async () => {
    try {
      const res = await fetch('/sample.csv')
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      // Unticked metrics carry 0 and metrics below the quorum are left out, so the sample passes the check as it is
      const text = rewriteSampleMonths(await res.text(), months)
        .split('\n')
        .flatMap((line, i) => {
          if (i === 0 || line.trim() === '') return [line]
          const cols = line.split(',')
          const rule = valueRule(quorumRules, cols[1], cols[2])
          if (rule === 'ignored') return []
          return [rule === 'zero' ? [...cols.slice(0, 4), '0'].join(',') : line]
        })
        .join('\n')
      download(`sample-${months[0]}.csv`, text)
    } catch (e) {
      console.error('[epoch] sample download failed', e)
    }
  }

  // After a failed submit, the page asks the server which values it holds
  const checkStored = async (message) => {
    try {
      const s = await loadSubmissions()
      const n = s.keys.size
      if (n >= owed) setSubmit({ phase: 'idle' })
      else setSubmit({ phase: 'failed', kind: n === 0 ? 'none' : 'part', stored: n, message })
    } catch {
      setSubmit({ phase: 'failed', kind: 'unknown', message })
    }
  }

  const doSubmit = async () => {
    const rows = upload.result?.rows ?? []
    setSubmit({ phase: 'submitting' })
    try {
      const token = await api.acquireApiToken(instance, account)
      const done = submitted?.keys ?? new Set()
      const payload = []
      for (const r of rows) {
        if (done.has(cellKey(r.country, r.month, r.indicator, r.segment))) continue
        const masked = await calculateMaskedValue(r.value, r.country, r.month, r.indicator, r.segment, myId, kx.sharedSecrets)
        payload.push({ country: r.country, month: r.month, indicator: r.indicator, segment: r.segment, value: masked.toString(), epochId, signature: 'web-ui' })
      }
      if (payload.length > 0) await api.submitMetricsBatch(payload, token)
      await loadSubmissions()
      setSubmit({ phase: 'idle' })
      setUpload({ phase: 'idle' })
    } catch (e) {
      console.error('[epoch] submit failed', e)
      await checkStored([e.status, e.message].filter(Boolean).join(' '))
    }
  }

  // ── Which wireframe state the page is in ─────────────────────────────────────
  let strip
  let content
  if (epoch.cancelledAt) {
    strip = null
    content = (
      <Alert kind="error" title="This epoch was cancelled" lines={[
        `The platform administrator cancelled it on ${longDate(epoch.cancelledAt)}. No further key exchange or submissions are accepted for it.`,
        epoch.cancelReason && `Reason: ${epoch.cancelReason}`,
      ]}>
        {epoch.replacedByEpochId && <div className="actions"><Link className="btn btn-primary" to={`/${u}?epoch=${epoch.replacedByEpochId}`}>Continue in the new epoch</Link></div>}
      </Alert>
    )
  } else if (viewOnly) {
    strip = null
    content = <Alert kind="warning" title={isAdmin ? 'Admins can’t contribute data' : 'Your organization is not a partner in this epoch'} lines={[isAdmin ? 'Follow this epoch in Admin.' : 'You can see its status, but not exchange keys or submit data.']} />
  } else if (epoch.isClosed) {
    strip = [D, D, D, D]
    content = (
      <>
        <PartNow><p>Nothing to do. The {monthYear(epoch.startDate)} epoch is closed.{submitted?.at ? ` You submitted your contribution on ${longDate(submitted.at)}.` : ''}</p></PartNow>
        <Sec h="What happens next"><p>Only combined totals are published, outside this tool. Individual values are never published. When the next epoch opens, it appears in your epoch list on <Link className="link" to="/">Home</Link>.</p></Sec>
      </>
    )
  } else if (allSubmitted) {
    strip = [D, D, D, D]
    content = (
      <>
        <PartNow><p>Nothing to do. You submitted your contribution{submitted?.at ? ` on ${longDate(submitted.at)}` : ''}. Your browser masked your {submitted?.keys.size ?? owed} values before it sent them. This is final.</p></PartNow>
        <Sec h="What happens next"><p>The epoch closes when every partner has submitted. Only combined totals are published, outside this tool. You can close this page.</p></Sec>
      </>
    )
  } else if (!kx.ready) {
    const registeredOn = kx.keyPair?.registeredAt ? `Your key was registered on ${longDate(kx.keyPair.registeredAt)}` : 'Your key is registered'
    const leave = 'Key exchange runs only while this page is open. Nothing will notify you, so come back to this page later. Use this browser until the epoch ends. Your key exists only in this browser.'
    if (kx.error) {
      strip = kx.phase === 'creating' ? [P, L, L, L] : [D, P, L, L]
      content = (
        <PartNow>
          <Alert kind="error" title={kx.phase === 'creating' ? 'Your key was not registered' : 'Key exchange did not finish'} lines={[
            `Server message: ${kx.error}`,
            kx.phase === 'creating' ? 'Try again. If it fails again, contact the platform administrator.' : `${registeredOn} and is still in this browser. Try again. If it fails again, contact the platform administrator.`,
          ]} />
          <div className="actions"><button className="btn btn-primary" onClick={kx.retry}>Try again</button></div>
        </PartNow>
      )
    } else if (kx.phase === 'creating') {
      strip = [W, L, L, L]
      content = (
        <PartNow>
          <p>Your browser is creating your key and registering it. Keep this page open until this step is done.</p>
          <p className="leave">The private key stays in this browser. Only the public key is sent to the server. Use this browser until the epoch ends.</p>
        </PartNow>
      )
    } else if (kx.pollError) {
      strip = [D, W, L, L]
      content = (
        <PartNow>
          <Alert kind="warning" title="This page can't reach the server" lines={[`The status on this page may be out of date.${kx.lastPollAt ? ` Last update: ${clockTime(kx.lastPollAt)}.` : ''} The page tries again every 2 seconds.`, `${registeredOn} and is still in this browser.`]} />
          <p>Nothing to do now. Key exchange starts when every partner in the epoch has registered a key.</p>
          <p className="leave">{leave}</p>
        </PartNow>
      )
    } else if (kx.phase === 'waiting-keys') {
      strip = [D, W, L, L]
      content = (
        <PartNow>
          <p>{registeredOn}. Nothing to do now. Key exchange starts when every partner in the epoch has registered a key.</p>
          <p className="leave">{leave}</p>
        </PartNow>
      )
    } else {
      strip = [D, W, L, L]
      content = (
        <>
          <PartNow>
            <p>Your part of the key setup is running. Other partners are still deriving their shared secrets.</p>
            <p className="leave">Keep this page open until this step is done. Use this browser until the epoch ends. Your key exists only in this browser.</p>
          </PartNow>
          <Sec h={SEC_FILE}>
            <p>Checking opens when every partner in the epoch has finished step 2. You can prepare your file now. It needs {owed} rows: {COUNTRIES.length} countries, {CELL_COUNT / COUNTRIES.length / 4} series, months {monthRange(epoch.startDate, 4)}{belowQuorumText ? `, without the metrics below the quorum: ${belowQuorumText}` : ''}.</p>
            <div className="actions">
              <button className="btn btn-primary" disabled aria-describedby="why-check">Upload file</button>
              <button className="btn btn-secondary" onClick={downloadSample}>Download sample.csv</button>
              <span className="why" id="why-check">Available when all partners finish step 2</span>
            </div>
          </Sec>
        </>
      )
    }
  } else if (!quorum?.quorumComplete) {
    strip = [D, D, W, L]
    content = (
      <>
        <Sec h={SEC_FILE}>
          <div className="lead">
            <p>Upload opens when the Quorum Check for this period is complete. You can prepare your file now.</p>
            <Link className="link" to={`/quorum?type=${u}`}>Go to the Quorum Check</Link>
          </div>
          <div className="actions">
            <button className="btn btn-primary" disabled aria-describedby="why-quorum">Upload file</button>
            <button className="btn btn-secondary" onClick={downloadSample}>Download sample.csv</button>
            <span className="why" id="why-quorum">Available when the Quorum Check is complete</span>
          </div>
        </Sec>
        <SubmitDisabled />
      </>
    )
  } else if (submit.phase === 'submitting' || submit.phase === 'failed') {
    strip = submit.phase === 'submitting' ? [D, D, D, W] : [D, D, D, P]
    content = (
      <>
        <ValidFile upload={upload} busy={submit.phase === 'submitting'} onChoose={choose} />
        <Sec h={SEC_SUBMIT}>
          <p>{SUBMIT_P}</p>
          {submit.phase === 'submitting' ? (
            <div className="actions"><button className="btn btn-primary" disabled aria-describedby="why-busy">Submitting</button><span className="why" id="why-busy">Masking {upload.result.rows.length} values, then sending them. Keep this page open.</span></div>
          ) : submit.kind === 'none' ? (
            <>
              <Alert kind="error" title="Submit did not finish" lines={[`The server did not confirm your contribution. Server message: ${submit.message}. The page then checked the server: none of your values are stored.`, `Your file is still loaded and checked. Submit again. If it fails again, contact ${SUPPORT_EMAIL}.`]} />
              <div className="actions"><button className="btn btn-primary" onClick={doSubmit}>Submit again</button></div>
            </>
          ) : submit.kind === 'part' ? (
            <Alert kind="error" title="Only part of your contribution is stored" lines={[`The submit did not finish. The page then checked the server: it has ${submit.stored} of your ${owed} values. Server message: ${submit.message}.`, 'Contact the platform administrator. Tell them the epoch and the number of values stored. Your file is still loaded in this page.']} />
          ) : (
            <>
              <Alert kind="error" title="Submit did not finish" lines={[`The server did not confirm your contribution. Server message: ${submit.message}. The page then tried to check which of your values the server has, and that check failed too.`, 'Your file is still loaded and checked. Check again before you submit again. If the check fails again, contact the platform administrator.']} />
              <div className="actions"><button className="btn btn-primary" onClick={() => checkStored(submit.message)}>Check again</button></div>
            </>
          )}
        </Sec>
      </>
    )
  } else if (upload.phase === 'checked' && upload.result.ok) {
    strip = [D, D, D, C]
    content = (
      <>
        <ValidFile upload={upload} onChoose={choose} />
        <Sec h={SEC_SUBMIT}><p>{SUBMIT_P}</p><div className="actions"><button className="btn btn-primary" onClick={doSubmit}>Submit contribution</button></div></Sec>
      </>
    )
  } else if (upload.phase === 'checked') {
    strip = [D, D, P, L]
    content = (
      <>
        <FileCheck upload={upload} onChoose={choose} onSample={downloadSample} />
        <SubmitDisabled />
      </>
    )
  } else if (upload.phase === 'checking') {
    strip = [D, D, W, L]
    content = (
      <>
        <Sec h={SEC_FILE}>
          <p>Checking your file in this browser.</p>
          <div className="actions"><button className="btn btn-primary" disabled>Upload file</button><button className="btn btn-secondary" disabled>Download sample.csv</button></div>
        </Sec>
        <SubmitDisabled />
      </>
    )
  } else {
    strip = [D, D, C, L]
    content = (
      <>
        <Sec h={SEC_FILE}>
          <p>Expecting monthly token data for {COUNTRIES.length} countries, {ACCESS.length} model access types and {REGIONS.length} model regions of origin.{epoch.quorumPeriodId ? ' Report values for the metrics you ticked in the Quorum Check, and enter 0 for the others.' : ''}{belowQuorumText ? ` Leave out the metrics below the quorum: ${belowQuorumText}. The file needs ${owed} rows.` : ''}</p>
          <div className="actions"><button className="btn btn-primary" onClick={choose}>Upload file</button><button className="btn btn-secondary" onClick={downloadSample}>Download sample.csv</button></div>
        </Sec>
        <SubmitDisabled />
      </>
    )
  }

  // Announce the current step when it changes
  const stripKey = strip?.join(',')
  const lastStrip = useRef(null)
  useEffect(() => {
    if (!strip) return
    if (lastStrip.current !== null && lastStrip.current !== stripKey) announce(stepAnnouncement(strip))
    lastStrip.current = stripKey
  }, [stripKey])

  return (
    <EpochShell u={u} epoch={epoch}>
      {strip && <ProtocolStrip strip={strip} />}
      <div className="main">{content}</div>
      <input ref={fileInput} type="file" accept=".csv,text/csv" onChange={onFile} hidden />
    </EpochShell>
  )
}

function SubmitDisabled() {
  return (
    <Sec h={SEC_SUBMIT}>
      <p>{SUBMIT_P}</p>
      <div className="actions"><button className="btn btn-primary" disabled aria-describedby="why-submit">Submit contribution</button><span className="why" id="why-submit">Available when the file is valid</span></div>
    </Sec>
  )
}

function FileInfo({ upload, meta }) {
  return (
    <div className="file">
      <div className="fname">{upload.fileName}</div>
      <div className="fmeta">{meta ?? `${upload.result.dataRowCount} rows | Checked in this browser at ${clockTime(upload.at)}`}</div>
    </div>
  )
}

function ValidFile({ upload, busy = false, onChoose }) {
  const rows = upload.result.rows.filter((r) => r.value > 0n)
  const countries = new Set(rows.map((r) => r.country)).size
  const access = new Set(rows.filter((r) => r.indicator === 'token_access').map((r) => r.segment)).size
  const regions = new Set(rows.filter((r) => r.indicator === 'model_region_of_origin').map((r) => r.segment)).size
  return (
    <Sec h={SEC_FILE}>
      <FileInfo upload={upload} />
      <div className="summary"><span className="big">All {upload.result.dataRowCount} rows are valid</span></div>
      <p>Your contribution: {plural(countries, 'country', 'countries')}, {plural(access, 'model access type', 'model access types')} and {plural(regions, 'model region of origin', 'model regions of origin')}.</p>
      <div className="actions"><button className="btn btn-secondary" onClick={onChoose} disabled={busy}>Upload a different file</button></div>
    </Sec>
  )
}

function ExampleTable({ g, limit }) {
  const rows = limit ? g.rows.slice(0, limit) : g.rows
  return (
    <div className="table-scroll">
      <table className="ex">
        <thead><tr>{g.cols.map((h) => <th key={h} scope="col">{h}</th>)}</tr></thead>
        <tbody>{rows.map((r, i) => <tr key={i}>{r.map((v, j) => <td key={j}>{j === g.bad ? <mark>{v}</mark> : v}</td>)}</tr>)}</tbody>
      </table>
    </div>
  )
}

function ProblemGroup({ g, i, compact }) {
  const [all, setAll] = useState(false)
  const more = g.rows.length > 3
  const copyHeader = () => navigator.clipboard?.writeText(HEADERS.join(',')).then(() => announce('Accepted header copied.'), () => {})
  return (
    <li className="prob">
      <span className="pn">{i + 1}</span>
      <div className="pc">
        <p className="what">{g.what}</p>
        <p className="where">{g.where}</p>
        {!compact && g.accepted && <p>{g.accepted}</p>}
        {(g.code ?? []).map(([l, v]) => [<p key={`${l}l`}>{l}</p>, <p key={`${l}v`} className="code">{v}</p>])}
        {g.copyHeader && <div><button className="btn btn-secondary" onClick={copyHeader}>Copy accepted header</button></div>}
        <p>{g.fix}</p>
        {g.rows.length > 0 && (compact
          ? <details className="exrows"><summary className="link">Show example rows</summary><ExampleTable g={g} limit={3} /></details>
          : <ExampleTable g={g} limit={all ? 0 : 3} />)}
        {!compact && more && <div><button className="link" onClick={() => setAll(!all)}>{all ? 'Show fewer rows' : `Show all ${g.rows.length} rows`}</button></div>}
      </div>
    </li>
  )
}

function FileCheck({ upload, onChoose, onSample }) {
  const v = upload.view
  const acts = v.acts ?? [['choose', 'Upload corrected file'], ['report', 'Download problem report (.csv)']]
  const run = {
    choose: onChoose,
    sample: onSample,
    report: () => download('problem-report.csv', problemReportCsv(v)),
    reload: () => window.location.reload(),
  }
  return (
    <Sec h={SEC_FILE}>
      <FileInfo upload={upload} meta={v.blocked ? v.meta : undefined} />
      <div className="summary">
        <div className="line">{v.blocked ? <Tag kind="blocked">Blocked</Tag> : <Tag kind="fix">Needs correction</Tag>}<span className="big">{v.big}</span></div>
        {v.sub && <p className="subl">{v.sub}</p>}
      </div>
      {(v.text ?? []).map((p, i) => <p key={i}>{p}</p>)}
      {v.groups?.length > 0 && <ol className="problems">{v.groups.map((g, i) => <ProblemGroup key={i} g={g} i={i} compact={v.compact} />)}</ol>}
      <div className="actions">
        {acts.map(([a, label], i) => <button key={a} className={`btn ${i ? 'btn-secondary' : 'btn-primary'}`} onClick={run[a]}>{label}</button>)}
      </div>
    </Sec>
  )
}
