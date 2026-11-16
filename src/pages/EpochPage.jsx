import { useCallback, useEffect, useRef, useState } from 'react'
import { Link, useSearchParams } from 'react-router-dom'
import { useMsal } from '@azure/msal-react'
import { Page, announce, download, useShell } from '../components/Shell'
import { Alert, ProtocolStrip, epochTag, stepAnnouncement } from '../components/ui'
import { FileCheck, FileInfo, SEC_FILE, Sec, ValidFile } from '../components/FileCheckView'
import useEpochs, { byStartDesc } from '../hooks/useEpochs'
import useKeyExchange from '../hooks/useKeyExchange'
import * as api from '../utils/api'
import { calculateMaskedValue } from '../utils/noise'
import { getDeviceId } from '../utils/localCrypto'
import {
  COUNTRIES, SEGMENTS_BY_INDICATOR, SERIES, cellCount, cellKey, epochMonths, isCsvFile, monthCount, pairKey,
  parseAndValidateCsv, rewriteSampleMonths, valueRule,
} from '../utils/csvUpload'
import { buildFileCheck } from '../utils/fileCheck'
import { answersMap, loadQuorumLocal } from '../utils/quorumLocal'
import { MIN_PASSPHRASE, OFFLINE_SIGNATURE, OFFLINE_TOOL_FILE, buildKeyFile, checkMaskedFile } from '../utils/offlineBundle'
import { USAGE, SUPPORT_EMAIL, clockTime, longDate, monthRange, monthYear } from '../utils/usage'

const POLL_MS = 5000
const D = 'done', C = 'current', W = 'waiting', P = 'problem', L = 'later'
const SEC_SUBMIT = 'Masked data submission'
const SUBMIT_P = 'Your browser masks each value, then sends only the masked values. Submitting is final: you cannot change your contribution after this.'
const SUBMIT_OFFLINE_P = 'These values were masked on your offline computer. This page sends them as they are. Submitting is final: you cannot change your contribution after this.'
const SEC_OFFLINE = 'No internet on the computer with your data?'
const ACCESS = SEGMENTS_BY_INDICATOR.token_access
const REGIONS = SEGMENTS_BY_INDICATOR.model_region_of_origin

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
        {epoch && <p className="meta">Reporting months: {monthRange(epoch.startDate, monthCount(epoch))}</p>}
      </div>
      {children}
    </main>
  )
}

const PartNow = ({ children }) => <Sec h="Your part now">{children}</Sec>

function EpochView({ u, epoch, isAdmin }) {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  const myId = account?.localAccountId ?? ''
  const deviceId = getDeviceId()
  const epochId = epoch.epochId
  const months = epochMonths(epoch) ?? []
  const maskBits = epoch.maskBits ?? 64
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
  const owed = cellCount(months.length) - ignoredCellCount
  const belowQuorumText = [...quorumRules.ignoredPairs].map((k) => k.replace('|', ' ')).join(', ')
  const allSubmitted = epoch.mySubmissionComplete || (submitted && submitted.keys.size >= owed && owed > 0)

  // ── Upload and submit ────────────────────────────────────────────────────────
  const [upload, setUpload] = useState({ phase: 'idle' })   // idle | checking | checked | masked-invalid
  const [submit, setSubmit] = useState({ phase: 'idle' })   // idle | submitting | failed
  const fileInput = useRef(null)
  const maskedInput = useRef(null)

  const choose = () => fileInput.current?.click()
  const chooseMasked = () => maskedInput.current?.click()

  // ── Offline submission: export a key file, submit the masked file made offline ─
  const [keyExport, setKeyExport] = useState({ phase: 'idle' })   // idle | exporting | done | failed

  const exportKeyFile = async (passphrase) => {
    setKeyExport({ phase: 'exporting' })
    try {
      const file = await buildKeyFile({
        epoch, months, maskBits, myId, keyPair: kx.keyPair, sharedSecrets: kx.sharedSecrets, quorumRules, passphrase,
        epochLabel: `${monthYear(epoch.startDate)} ${USAGE[u].label}`,
      })
      const name = `airt-epoch-${epochId}-key.json`
      download(name, JSON.stringify(file), 'application/json')
      setKeyExport({ phase: 'done', name })
      announce(`Key file ${name} saved.`)
    } catch (e) {
      console.error('[epoch] key file export failed', e)
      setKeyExport({ phase: 'failed', message: e.message })
    }
  }

  const onMaskedFile = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    setSubmit({ phase: 'idle' })
    setUpload({ phase: 'checking', fileName: file.name })
    const at = new Date()
    let check
    try {
      check = await checkMaskedFile(await file.text(), { epochId, myId, maskBits, keyPair: kx.keyPair, months, quorumRules })
    } catch (e) {
      console.error('[epoch] masked file check failed', e)
      check = { ok: false, problems: ['The file could not be read. Choose the .json file the offline tool created.'] }
    }
    if (check.ok) {
      setUpload({ phase: 'checked', offline: true, fileName: file.name, at, result: { ok: true, rows: check.rows, dataRowCount: check.rows.length, maskedFile: check.file } })
    } else {
      setUpload({ phase: 'masked-invalid', fileName: file.name, at, problems: check.problems })
    }
  }

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
        try { result = parseAndValidateCsv(text, months, quorumRules, maskBits) } catch (e) {
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
        const masked = upload.offline ? r.masked : await calculateMaskedValue(r.value, r.country, r.month, r.indicator, r.segment, myId, kx.sharedSecrets, maskBits)
        payload.push({ country: r.country, month: r.month, indicator: r.indicator, segment: r.segment, value: masked.toString(), epochId, signature: upload.offline ? OFFLINE_SIGNATURE : 'web-ui' })
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
            <p>Checking opens when every partner in the epoch has finished step 2. You can prepare your file now. It needs {owed} rows: {COUNTRIES.length} countries, {SERIES.length} series, months {monthRange(epoch.startDate, monthCount(epoch))}{belowQuorumText ? `, without the metrics below the quorum: ${belowQuorumText}` : ''}.</p>
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
            <Link className="link" to="/quorum">Go to the Quorum Check</Link>
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
        {upload.offline
          ? <MaskedFile upload={upload} busy={submit.phase === 'submitting'} onChoose={chooseMasked} />
          : <ValidFile upload={upload} busy={submit.phase === 'submitting'} onChoose={choose} />}
        <Sec h={SEC_SUBMIT}>
          <p>{upload.offline ? SUBMIT_OFFLINE_P : SUBMIT_P}</p>
          {submit.phase === 'submitting' ? (
            <div className="actions"><button className="btn btn-primary" disabled aria-describedby="why-busy">Submitting</button><span className="why" id="why-busy">{upload.offline ? `Sending ${upload.result.rows.length} masked values` : `Masking ${upload.result.rows.length} values, then sending them`}. Keep this page open.</span></div>
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
        {upload.offline ? <MaskedFile upload={upload} onChoose={chooseMasked} /> : <ValidFile upload={upload} onChoose={choose} />}
        <Sec h={SEC_SUBMIT}><p>{upload.offline ? SUBMIT_OFFLINE_P : SUBMIT_P}</p><div className="actions"><button className="btn btn-primary" onClick={doSubmit}>Submit contribution</button></div></Sec>
      </>
    )
  } else if (upload.phase === 'masked-invalid') {
    strip = [D, D, P, L]
    content = (
      <>
        <Sec h={SEC_FILE}>
          <FileInfo upload={upload} meta={`Masked file | Checked in this browser at ${clockTime(upload.at)}`} />
          <Alert kind="error" title="This masked file can't be submitted" lines={upload.problems} />
          <div className="actions">
            <button className="btn btn-primary" onClick={chooseMasked}>Choose a different masked file</button>
            <button className="btn btn-secondary" onClick={() => setUpload({ phase: 'idle' })}>Back</button>
          </div>
        </Sec>
        <SubmitDisabled />
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
        <OfflineSubmission keyExport={keyExport} onExport={exportKeyFile} onChooseMasked={chooseMasked} />
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
      <input ref={maskedInput} type="file" accept=".json,application/json" onChange={onMaskedFile} hidden />
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

function MaskedFile({ upload, busy = false, onChoose }) {
  const f = upload.result.maskedFile
  return (
    <Sec h={SEC_FILE}>
      <FileInfo upload={upload} meta={`Masked offline on ${longDate(f.createdAt)} | Checked in this browser at ${clockTime(upload.at)}`} />
      <div className="summary"><span className="big">All {upload.result.rows.length} masked values are valid</span></div>
      <p>The file matches this epoch and the key in this browser. Masked values can't be checked against your CSV here; the offline tool checked the CSV before it masked it.</p>
      <div className="actions"><button className="btn btn-secondary" onClick={onChoose} disabled={busy}>Choose a different masked file</button></div>
    </Sec>
  )
}

function OfflineSubmission({ keyExport, onExport, onChooseMasked }) {
  const [passphrase, setPassphrase] = useState('')
  const [confirm, setConfirm] = useState('')
  const busy = keyExport.phase === 'exporting'
  const tooShort = passphrase.length < MIN_PASSPHRASE
  const mismatch = confirm !== '' && confirm !== passphrase
  const submitExport = (event) => {
    event.preventDefault()
    if (tooShort || confirm !== passphrase) return
    onExport(passphrase)
    setPassphrase('')
    setConfirm('')
  }
  return (
    <Sec h={SEC_OFFLINE}>
      <p>You can mask your data on a computer without internet access:</p>
      <ol className="offline-steps">
        <li>Export your key file here, protected with a passphrase, and download the offline tool.</li>
        <li>Copy both files to the computer with your data. Open {OFFLINE_TOOL_FILE} in Edge, Chrome or Firefox, open the key file, upload your CSV and save the masked file.</li>
        <li>Copy the masked file back to this computer and submit it here, in this browser.</li>
      </ol>
      <p className="leave">Anyone with the key file and its passphrase can unmask your values. Don’t send the passphrase with the file, and delete the key file from both computers after you submit.</p>
      {keyExport.phase === 'failed' && <Alert kind="error" title="The key file was not exported" lines={[keyExport.message, `Try again. If it fails again, contact ${SUPPORT_EMAIL}.`]} />}
      {keyExport.phase === 'done' && <Alert kind="success" title={`Saved ${keyExport.name}`} lines={['Check your downloads folder.']} />}
      <form className="offline-form" onSubmit={submitExport}>
        <label className="field">Passphrase for the key file (at least {MIN_PASSPHRASE} characters)
          <input type="password" autoComplete="new-password" value={passphrase} onChange={(e) => setPassphrase(e.target.value)} disabled={busy} />
        </label>
        <label className="field">Repeat the passphrase
          <input type="password" autoComplete="new-password" value={confirm} onChange={(e) => setConfirm(e.target.value)} disabled={busy} aria-invalid={mismatch} aria-describedby={mismatch ? 'pp-mismatch' : undefined} />
        </label>
        {mismatch && <span className="why" id="pp-mismatch">The passphrases don’t match</span>}
        <div className="actions">
          <button className="btn btn-primary" type="submit" disabled={busy || tooShort || confirm !== passphrase}>{busy ? 'Exporting' : 'Export key file'}</button>
          <a className="btn btn-secondary" href={`/${OFFLINE_TOOL_FILE}`} download={OFFLINE_TOOL_FILE}>Download offline tool</a>
          <button className="btn btn-secondary" type="button" onClick={onChooseMasked}>Submit masked file</button>
        </div>
      </form>
    </Sec>
  )
}
