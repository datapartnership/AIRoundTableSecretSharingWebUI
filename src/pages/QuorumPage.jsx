import { useEffect, useRef, useState } from 'react'
import { useMsal } from '@azure/msal-react'
import { Page, announce, download, useShell } from '../components/Shell'
import { Alert, Checkbox, Tag } from '../components/ui'
import useKeyExchange from '../hooks/useKeyExchange'
import * as api from '../utils/api'
import { calculateMaskedValue } from '../utils/noise'
import { getDeviceId } from '../utils/localCrypto'
import { COUNTRIES, INDICATORS, pairKey, parseQuorumCsv, quorumSampleCsv } from '../utils/csvUpload'
import { loadQuorumLocal, saveQuorumLocal } from '../utils/quorumLocal'
import { USAGE, SUPPORT_EMAIL, clockTime, longDate, monthYear } from '../utils/usage'

const COUNTRY_NAMES = { IND: 'India', BRA: 'Brazil', USA: 'United States', GBR: 'United Kingdom', NGA: 'Nigeria' }
const INDICATOR_LABELS = { tokens: 'Tokens', token_access: 'Access type', model_region_of_origin: 'Model region of origin' }
const NI = INDICATORS.length
const NQ = COUNTRIES.length * NI
const cname = (c) => `${COUNTRY_NAMES[c] ?? c} (${c})`
const keyAt = (i) => pairKey(COUNTRIES[Math.floor(i / NI)], INDICATORS[i % NI])
const count = (a) => a.filter(Boolean).length

export default function QuorumPage() {
  // The one active Quorum Check, polled by the app shell so every page sees the same period
  const { isAdmin, quorum } = useShell()
  const { status, known, error: loadError, refresh } = quorum
  const period = status?.period ?? null
  const u = period?.usageType ?? null

  // An admin cancelled the period and opened another while this page was open
  const seenPeriod = useRef(null)
  const [replaced, setReplaced] = useState(false)
  useEffect(() => {
    if (!period) return
    if (seenPeriod.current !== null && seenPeriod.current !== period.periodId) {
      setReplaced(true)
      announce('The platform administrator opened a new Quorum Check.')
    }
    seenPeriod.current = period.periodId
  }, [period?.periodId])

  const tag = !period ? null : period.isCancelled ? <Tag kind="blocked">Cancelled</Tag>
    : status.quorumComplete ? <Tag kind="closed">Complete</Tag> : <Tag kind="open">Open</Tag>

  let body
  if (loadError && !status) {
    body = <Alert kind="error" title="The Quorum Check could not be loaded" lines={[`Server message: ${loadError}`, 'Your answers are not affected. Reload the page, or contact the platform administrator if this continues.']} />
  } else if (!known || !status) {
    body = <div role="status"><p className="qsum">Checking the Quorum Check</p></div>
  } else if (!period) {
    body = <section className="qsec"><h2>No Quorum Check is open</h2><p>It appears here when the platform administrator opens one. This page updates automatically.</p></section>
  } else if (isAdmin || !status.isEligible) {
    body = <section className="qsec"><h2>Your answers</h2><p className="note">{isAdmin ? 'Admins can’t answer the Quorum Check. Partners’ progress is shown in Admin.' : 'Your organization is not a partner in this Quorum Check.'}</p></section>
  } else {
    body = (
      <>
        {replaced && <Alert kind="warning" title="A new Quorum Check was opened" lines={['The platform administrator replaced the previous Quorum Check. Answers given to it no longer count, so answer this one.']} />}
        <QuorumAnswers key={period.periodId} u={u} status={status} onSubmitted={refresh} />
      </>
    )
  }

  return (
    <Page>
      <main id="main" className="epoch quorum wrap">
        <div className="epoch-head">
          <div className="title-row"><h1 tabIndex={-1}>Quorum check</h1>{tag}</div>
          {period && <p className="meta">Covers {USAGE[u].label} epochs from {monthYear(period.startDate)} to {monthYear(period.startDate, 11)}</p>}
        </div>
        <div className="main">{body}</div>
      </main>
    </Page>
  )
}

function QuorumAnswers({ u, status, onSubmitted }) {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  const deviceId = getDeviceId()
  const period = status.period
  const myId = account?.localAccountId ?? ''
  const local = loadQuorumLocal(myId, period.periodId, deviceId)

  const kx = useKeyExchange({
    scope: { periodId: period.periodId },
    producerIds: period.producerIds,
    active: !status.myDone && !period.isCancelled,
  })

  const [answers, setAnswers] = useState(() => Array(NQ).fill(false))
  const [csv, setCsv] = useState(null)          // { name } once a CSV filled the boxes
  const [csvError, setCsvError] = useState(null)
  const [busy, setBusy] = useState(false)
  const [submitError, setSubmitError] = useState(null)
  const dialog = useRef(null)
  const fileInput = useRef(null)
  const submitBtn = useRef(null)
  const n = count(answers)
  const range = `${monthYear(period.startDate)} to ${monthYear(period.startDate, 11)}`

  const setOne = (i, v) => setAnswers((a) => a.map((x, j) => (j === i ? v : x)))
  const setRow = (r, v) => setAnswers((a) => a.map((x, j) => (Math.floor(j / NI) === r ? v : x)))
  const setCol = (c, v) => setAnswers((a) => a.map((x, j) => (j % NI === c ? v : x)))

  const onCsv = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    setCsvError(null)
    try {
      const parsed = parseQuorumCsv(await file.text())
      if (parsed.error) { setCsvError(parsed.error); return }
      setAnswers(Array.from({ length: NQ }, (_, i) => parsed.answers.get(keyAt(i)) === true))
      setCsv({ name: file.name })
      announce(`${NQ} answers loaded into the boxes. Nothing is submitted yet.`)
    } catch {
      setCsvError('This browser couldn’t open the file. Close it in any other program, then choose it again.')
    }
  }

  const submit = async () => {
    dialog.current?.close()
    setBusy(true)
    setSubmitError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const randomTag = () => {
        const b = crypto.getRandomValues(new BigUint64Array(1))[0]
        return b === 0n ? 1n : b
      }
      const tagA = local?.tagA ? BigInt(local.tagA) : randomTag()
      const tagB = local?.tagB ? BigInt(local.tagB) : randomTag()
      const month = String(period.startDate).slice(0, 7)
      const stored = {}
      const rows = []
      for (let i = 0; i < NQ; i++) {
        const country = COUNTRIES[Math.floor(i / NI)]
        const indicator = INDICATORS[i % NI]
        const flag = answers[i] ? 1n : 0n
        stored[pairKey(country, indicator)] = Number(flag)
        const mask = (lane, v) => calculateMaskedValue(v, country, month, `quorum-${lane}:${indicator}`, 'all', myId, kx.sharedSecrets)
        rows.push({
          country,
          indicator,
          maskedCount: (await mask('count', flag)).toString(),
          maskedTagA: (await mask('tagA', flag * BigInt.asIntN(64, tagA))).toString(),
          maskedTagB: (await mask('tagB', flag * BigInt.asIntN(64, tagB))).toString(),
        })
      }
      // Persist before sending so the answers survive a failed or interrupted request
      saveQuorumLocal(myId, period.periodId, deviceId, { tagA: tagA.toString(), tagB: tagB.toString(), answers: stored, submittedAt: new Date().toISOString() })
      await api.submitQuorum(period.periodId, rows, token)
      announce('Your answers were submitted.')
      onSubmitted()
    } catch (e) {
      console.error('[quorum] submit failed', e)
      setSubmitError(e.message)
    } finally {
      setBusy(false)
    }
  }

  if (status.myDone) {
    const saved = local?.answers ?? null
    const at = new Date(status.myAnsweredAt ?? local?.submittedAt ?? Date.now())
    return (
      <section className="qsec">
        <h2>Your answers</h2>
        <p>
          Your answers were submitted on {longDate(at)} at {clockTime(at)}.{' '}
          {status.quorumComplete
            ? 'Every partner has answered, so the data upload is open.'
            : `${status.answeredPartners} of ${status.partnerCount} partners have answered. This page updates automatically and takes you to data submission when everyone has answered.`}
        </p>
        {saved
          ? <Matrix answers={Array.from({ length: NQ }, (_, i) => saved[keyAt(i)] === 1)} done />
          : <p className="note">Your answers were given in another browser, so they can&apos;t be shown here.</p>}
        <p className="note">Your answers cannot change until the next Quorum Check.</p>
      </section>
    )
  }

  const keyIssue = kx.error
    ? <Alert kind="error" title="Key exchange did not finish" lines={[`Server message: ${kx.error}`, 'Your answers are still ticked on this page. Try again. If it fails again, contact the platform administrator.']}>
        <div className="actions"><button className="btn btn-primary" onClick={kx.retry}>Try again</button></div>
      </Alert>
    : kx.pollError
      ? <Alert kind="warning" title="This page can't reach the server" lines={[`The status on this page may be out of date.${kx.lastPollAt ? ` Last update: ${clockTime(kx.lastPollAt)}.` : ''} The page tries again every 2 seconds.`]} />
      : null

  return (
    <>
      <section className="qsec">
        <h2>Your answers</h2>
        <p>Tick each country and indicator your organization will report, or upload a CSV that ticks the boxes for you. Your answers apply to every epoch in this period. A metric is published only if at least 3 partners report it.</p>
        <div className="actions">
          <button className="btn btn-primary" onClick={() => fileInput.current?.click()}>{csv ? 'Upload a different CSV' : 'Upload CSV'}</button>
          <button className="btn btn-secondary" onClick={() => download('sample.csv', quorumSampleCsv())}>Download sample.csv</button>
          <input ref={fileInput} type="file" accept=".csv,text/csv" onChange={onCsv} hidden />
        </div>
        {csvError && <Alert kind="error" title="The CSV did not fill the boxes" lines={[csvError, 'Fix the file, then upload it again, or tick the boxes below.']} />}
        {csv && (
          <div className="file">
            <div className="fname">{csv.name}</div>
            <div className="fmeta">{NQ} answers loaded into the boxes below. Nothing is submitted yet. Check the boxes, then submit your answers.</div>
          </div>
        )}
        <Matrix answers={answers} setOne={setOne} setRow={setRow} setCol={setCol} />
      </section>
      <section className="qsec">
        <p className="qsum" aria-live="polite">You will report {n} of {NQ} metrics.</p>
        {keyIssue}
        {submitError && <Alert kind="error" title="Your answers were not submitted" lines={[`The server did not confirm your answers. Server message: ${submitError}.`, `Your answers are still ticked on this page. Submit again. If it fails again, contact ${SUPPORT_EMAIL}.`]} />}
        <div className="actions">
          <button ref={submitBtn} className="btn btn-primary" disabled={!kx.ready || busy} aria-describedby={kx.ready ? undefined : 'why-q'}
            onClick={() => dialog.current?.showModal()}>
            {busy ? 'Submitting' : submitError ? 'Submit again' : 'Submit answers'}
          </button>
          {!kx.ready && <span className="why" id="why-q">Available when every partner has opened this Quorum Check. Keep this page open.</span>}
        </div>
        <p className="note">Your answers are masked in your browser. No single party can see which metrics your organization reports. Your answers cannot change until the next Quorum Check.</p>
      </section>
      <dialog ref={dialog} className="dialog" aria-labelledby="q-dh">
        <div>
          <h2 id="q-dh">Submit your answers</h2>
          <p>You will report {n} of {NQ} metrics for {USAGE[u].label}, {range}.</p>
          <p>Your answers cannot change until the next Quorum Check.</p>
          <p>In every epoch of this period, report a value above 0 for each metric you ticked, and enter 0 for each metric you did not tick.</p>
          <div className="actions">
            <button className="btn btn-primary" onClick={submit}>Submit answers</button>
            <button className="btn btn-secondary" onClick={() => { dialog.current?.close(); submitBtn.current?.focus() }}>Go back</button>
          </div>
        </div>
      </dialog>
    </>
  )
}

// Row and column boxes are select-all controls: checked when all their boxes are ticked, mixed when some are.
function Matrix({ answers, done = false, setOne, setRow, setCol }) {
  const part = (pred) => answers.filter((_, i) => pred(i))
  const all = (a) => count(a) === a.length
  const mixed = (a) => count(a) > 0 && count(a) < a.length
  return (
    <div className="matrix-wrap">
      <table className="matrix">
        <thead>
          <tr>
            <th scope="col">Country</th>
            {INDICATORS.map((ind, c) => {
              const col = part((i) => i % NI === c)
              return (
                <th key={ind} scope="col">
                  <div className="colh">
                    {done ? <span>{INDICATOR_LABELS[ind]}</span> : <Checkbox checked={all(col)} indeterminate={mixed(col)} onChange={(v) => setCol(c, v)} label={INDICATOR_LABELS[ind]} />}
                    <span className="csvname">{ind}</span>
                  </div>
                </th>
              )
            })}
          </tr>
        </thead>
        <tbody>
          {COUNTRIES.map((ct, r) => {
            const row = part((i) => Math.floor(i / NI) === r)
            return (
              <tr key={ct}>
                <th scope="row">{done ? cname(ct) : <Checkbox checked={all(row)} indeterminate={mixed(row)} onChange={(v) => setRow(r, v)} label={cname(ct)} />}</th>
                {INDICATORS.map((ind, c) => {
                  const i = r * NI + c
                  return (
                    <td key={ind}>
                      {done ? (answers[i] ? 'Yes' : 'No') : <Checkbox checked={answers[i]} onChange={(v) => setOne(i, v)} ariaLabel={`${cname(ct)}, ${INDICATOR_LABELS[ind]}`} />}
                    </td>
                  )
                })}
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}
