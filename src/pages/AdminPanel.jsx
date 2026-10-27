import { useState, useEffect, useCallback } from 'react'
import { useMsal } from '@azure/msal-react'
import { Page } from '../components/Shell'
import { Alert, Checkbox, Tag } from '../components/ui'
import * as api from '../utils/api'
import { CELL_COUNT, MONTH_COUNT, SEGMENTS_BY_INDICATOR, formatInt } from '../utils/csvUpload'
import { USAGE, USAGE_TYPES, monthYear, monthRange, longDate } from '../utils/usage'

// Must match EpochLifecycle.MinParticipants in the API
const MIN_PARTICIPANTS = 3
const PERIOD_MONTHS = 12

const thisMonth = () => {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
}

const monthIndex = (ym) => {
  const m = /^(\d{4})-(\d{2})/.exec(String(ym ?? ''))
  return m ? Number(m[1]) * 12 + Number(m[2]) - 1 : null
}

// The non-cancelled period of the usage type whose 12 months hold the epoch's 4 months
const periodFor = (periods, usageType, startMonth) => {
  const s = monthIndex(startMonth)
  return periods.find((p) => p.usageType === usageType && !p.cancelledAt
    && s >= monthIndex(p.startDate) && s + 4 <= monthIndex(p.startDate) + PERIOD_MONTHS) ?? null
}

function formatMissingCells(cells, ignoredCount = 0) {
  const n = cells?.length ?? 0
  if (n === 0) return '—'
  const preview = cells.slice(0, 6).map((c) => `${c.country} ${c.month} ${c.indicator}/${c.segment}`)
  // Below-quorum cells get no data, so they are not required
  return `${n} of ${CELL_COUNT - ignoredCount} missing (${preview.join(', ')}${n > 6 ? ' …' : ''})`
}

// Cells of the metrics that fell below the quorum: every segment of the indicator, in every month
const ignoredCellCount = (detail) => (detail.quorumCells ?? [])
  .filter((c) => c.ignored)
  .reduce((n, c) => n + MONTH_COUNT * (SEGMENTS_BY_INDICATOR[c.indicator]?.length ?? 0), 0)

const statusTag = (x) => x.cancelledAt ? <Tag kind="blocked">Cancelled</Tag>
  : x.isClosed ? <Tag kind="closed">Closed</Tag>
  : x.quorumComplete === false ? <Tag kind="waiting">Quorum Check</Tag>
  : <Tag kind="open">Open</Tag>

export default function AdminPanel() {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  const myId = account?.localAccountId ?? ''
  const token = useCallback(() => api.acquireApiToken(instance, account), [instance, account])

  // ── Registered partners ─────────────────────────────────────────────────────
  const [registered, setRegistered] = useState([])
  const [regLoading, setRegLoading] = useState(false)
  const [regError, setRegError] = useState(null)
  const [selected, setSelected] = useState(new Set())
  const [clearBusy, setClearBusy] = useState(false)
  const [clearConfirm, setClearConfirm] = useState(false)

  // ── Quorum Check periods ────────────────────────────────────────────────────
  const [periods, setPeriods] = useState([])
  const [periodUsage, setPeriodUsage] = useState('api')
  const [periodStart, setPeriodStart] = useState(thisMonth)
  const [periodBusy, setPeriodBusy] = useState(false)
  const [periodMsg, setPeriodMsg] = useState(null)       // { kind, text }

  // ── Epoch creation ──────────────────────────────────────────────────────────
  const [epochUsage, setEpochUsage] = useState('api')
  const [startMonth, setStartMonth] = useState(thisMonth)
  const [epochBusy, setEpochBusy] = useState(false)
  const [epochMsg, setEpochMsg] = useState(null)
  const [confirmCreate, setConfirmCreate] = useState(false)

  // ── Epochs ──────────────────────────────────────────────────────────────────
  const [epochs, setEpochs] = useState([])
  const [epochsLoading, setEpochsLoading] = useState(false)
  const [epochsError, setEpochsError] = useState(null)
  const [selectedEpochId, setSelectedEpochId] = useState(null)
  const [epochDetail, setEpochDetail] = useState(null)
  const [detailLoading, setDetailLoading] = useState(false)
  const [detailError, setDetailError] = useState(null)

  // ── Cancel / recreate the selected epoch ────────────────────────────────────
  const [epochAction, setEpochAction] = useState(null) // 'cancel' | 'recreate' while confirming
  const [cancelReason, setCancelReason] = useState('')
  const [actionBusy, setActionBusy] = useState(false)
  const [actionMsg, setActionMsg] = useState(null)

  const loadRegistered = useCallback(async () => {
    setRegLoading(true)
    setRegError(null)
    try {
      const data = await api.getProducers(await token())
      setRegistered(data.filter((p) => p.producerId !== myId))
    } catch (e) {
      setRegError(e.message)
    } finally {
      setRegLoading(false)
    }
  }, [token, myId])

  const loadPeriods = useCallback(async () => {
    try {
      const data = await api.adminGetQuorumPeriods(await token())
      setPeriods(data.periods ?? [])
    } catch (e) {
      setPeriodMsg({ kind: 'error', text: e.message })
    }
  }, [token])

  const loadEpochs = useCallback(async () => {
    setEpochsLoading(true)
    setEpochsError(null)
    try {
      const data = await api.adminGetEpochs(await token())
      const list = data.epochs ?? []
      setEpochs(list)
      setSelectedEpochId((prev) => (prev && list.some((e) => e.epochId === prev) ? prev : list[0]?.epochId ?? null))
    } catch (e) {
      setEpochsError(e.message)
    } finally {
      setEpochsLoading(false)
    }
  }, [token])

  const loadEpochDetail = useCallback(async (epochId) => {
    if (!epochId) { setEpochDetail(null); return }
    setDetailLoading(true)
    setDetailError(null)
    try {
      setEpochDetail(await api.adminGetEpochDetail(epochId, await token()))
    } catch (e) {
      setEpochDetail(null)
      setDetailError(e.message)
    } finally {
      setDetailLoading(false)
    }
  }, [token])

  useEffect(() => {
    loadRegistered()
    loadPeriods()
    loadEpochs()
  }, [])

  useEffect(() => {
    loadEpochDetail(selectedEpochId)
    setEpochAction(null)
    setCancelReason('')
  }, [selectedEpochId, loadEpochDetail])

  const nameOf = (id) => registered.find((p) => p.producerId === id)?.displayName ?? id
  const asProducers = (ids) => ids.map((id) => ({ producerId: id, displayName: nameOf(id) }))

  const toggleSelect = (id, on) => setSelected((prev) => {
    const next = new Set(prev)
    if (on) next.add(id); else next.delete(id)
    return next
  })

  const handleCreatePeriod = async () => {
    setPeriodBusy(true)
    setPeriodMsg(null)
    try {
      const p = await api.adminCreateQuorumPeriod({ usageType: periodUsage, startMonth: periodStart, producers: asProducers([...selected]) }, await token())
      setPeriodMsg({ kind: 'ok', text: `Quorum Check period ${p.periodId} opened for ${USAGE[p.usageType].label}, ${monthYear(p.startDate)} to ${monthYear(p.startDate, PERIOD_MONTHS - 1)}, with ${p.producerCount} partners.` })
      setSelected(new Set())
      await loadPeriods()
    } catch (e) {
      setPeriodMsg({ kind: 'error', text: e.message })
    } finally {
      setPeriodBusy(false)
    }
  }

  const handleCancelPeriod = async (p) => {
    if (!window.confirm(`Cancel Quorum Check period ${p.periodId}? No new epochs can be created in it.`)) return
    try {
      await api.adminCancelQuorumPeriod(p.periodId, null, await token())
      await loadPeriods()
    } catch (e) {
      setPeriodMsg({ kind: 'error', text: e.message })
    }
  }

  const epochPeriod = periodFor(periods, epochUsage, startMonth)

  const handleCreateEpoch = async () => {
    setEpochBusy(true)
    setEpochMsg(null)
    try {
      const data = await api.adminResetAndCreateEpoch({ usageType: epochUsage, startMonth, producers: asProducers(epochPeriod.producerIds) }, await token())
      setEpochMsg({ kind: 'ok', text: `Epoch ${data.epoch?.epochId} created for ${USAGE[epochUsage].label}, ${monthYear(data.epoch?.startDate)}, with ${data.producers?.length} partners.` })
      await loadEpochs()
      if (data.epoch?.epochId) setSelectedEpochId(data.epoch.epochId)
    } catch (e) {
      setEpochMsg({ kind: 'error', text: e.message })
    } finally {
      setEpochBusy(false)
      setConfirmCreate(false)
    }
  }

  const handleEpochAction = async () => {
    const epochId = epochDetail?.epochId
    if (!epochId || !epochAction) return
    setActionBusy(true)
    setActionMsg(null)
    try {
      const reason = cancelReason.trim() || null
      if (epochAction === 'recreate') {
        const data = await api.adminRecreateEpoch(epochId, reason, await token())
        setActionMsg({ kind: 'ok', text: `Epoch ${epochId} was cancelled and replaced by epoch ${data.epoch?.epochId}.` })
        await loadEpochs()
        if (data.epoch?.epochId) setSelectedEpochId(data.epoch.epochId)
      } else {
        await api.adminCancelEpoch(epochId, reason, await token())
        setActionMsg({ kind: 'ok', text: `Epoch ${epochId} was cancelled.` })
        await loadEpochs()
        await loadEpochDetail(epochId)
      }
      setEpochAction(null)
      setCancelReason('')
    } catch (e) {
      setActionMsg({ kind: 'error', text: e.message })
    } finally {
      setActionBusy(false)
    }
  }

  const handleClearAll = async () => {
    setClearBusy(true)
    try {
      await api.adminReset(await token())
      setRegistered([])
      setSelected(new Set())
      setPeriods([])
      setEpochs([])
      setSelectedEpochId(null)
      setEpochDetail(null)
    } catch (e) {
      setRegError(e.message)
    } finally {
      setClearBusy(false)
      setClearConfirm(false)
    }
  }

  const allSelected = registered.length > 0 && selected.size === registered.length
  const missingCount = MIN_PARTICIPANTS - selected.size
  const msg = (m) => m && (m.kind === 'ok'
    ? <p className="ok" role="status">{m.text}</p>
    : <Alert kind="error" title="That did not work" lines={[`Server message: ${m.text}`]} />)

  return (
    <Page>
      <main id="main" className="admin wrap">
        <div className="epoch-head">
          <p className="usage">Platform administration</p>
          <h1 tabIndex={-1}>Admin</h1>
          <p className="meta">Open Quorum Check periods and epochs, and follow submissions.</p>
        </div>

        <section className="sec">
          <h2>Registered partners</h2>
          <p className="muted">Partners appear here automatically when they sign in. Select partners to open a Quorum Check period.</p>
          {regError && <Alert kind="error" title="Partners could not be loaded" lines={[`Server message: ${regError}`]} />}
          {registered.length === 0 && !regLoading ? (
            <p className="muted">No partners registered yet.</p>
          ) : (
            <div className="atable-wrap">
              <table className="atable">
                <thead><tr><th style={{ width: 40 }}><Checkbox checked={allSelected} indeterminate={selected.size > 0 && !allSelected} ariaLabel="Select all partners"
                  onChange={(v) => setSelected(v ? new Set(registered.map((p) => p.producerId)) : new Set())} /></th><th>Name</th><th>Object ID (producer ID)</th><th>Joined</th></tr></thead>
                <tbody>
                  {registered.map((p) => (
                    <tr key={p.producerId}>
                      <td><Checkbox checked={selected.has(p.producerId)} onChange={(v) => toggleSelect(p.producerId, v)} ariaLabel={`Select ${p.displayName}`} /></td>
                      <td><b>{p.displayName}</b></td>
                      <td className="mono">{p.producerId}</td>
                      <td>{p.joinedDate ? longDate(p.joinedDate) : '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <div className="actions">
            <button className="btn btn-secondary" onClick={loadRegistered} disabled={regLoading}>{regLoading ? 'Refreshing' : 'Refresh'}</button>
            {!clearConfirm
              ? <button className="btn btn-secondary" onClick={() => setClearConfirm(true)}>Clear all data</button>
              : <>
                  <span className="err">Wipes all partners, periods, epochs, keys, ciphertexts, answers and submissions.</span>
                  <button className="btn btn-primary" onClick={handleClearAll} disabled={clearBusy}>{clearBusy ? 'Clearing' : 'Confirm clear'}</button>
                  <button className="btn btn-secondary" onClick={() => setClearConfirm(false)}>Cancel</button>
                </>}
          </div>
        </section>

        <section className="sec">
          <h2>Quorum Check periods</h2>
          <p className="muted">A period covers {PERIOD_MONTHS} months of epochs of one usage type. Its partners answer the Quorum Check once, and every epoch in it has the same partners.</p>
          <div className="row">
            <label className="field">Usage type
              <select value={periodUsage} onChange={(e) => setPeriodUsage(e.target.value)}>
                {USAGE_TYPES.map((u) => <option key={u} value={u}>{USAGE[u].label}</option>)}
              </select>
            </label>
            <label className="field">Start month<input type="month" value={periodStart} onChange={(e) => setPeriodStart(e.target.value)} /></label>
            <button className="btn btn-primary" onClick={handleCreatePeriod} disabled={periodBusy || selected.size < MIN_PARTICIPANTS}>
              {periodBusy ? 'Opening' : `Open period with ${selected.size} selected partners`}
            </button>
          </div>
          {selected.size < MIN_PARTICIPANTS && <p className="why">Select {missingCount} more {missingCount === 1 ? 'partner' : 'partners'} above. A period needs at least {MIN_PARTICIPANTS}.</p>}
          {msg(periodMsg)}
          {periods.length > 0 && (
            <div className="atable-wrap">
              <table className="atable">
                <thead><tr><th>Period</th><th>Usage type</th><th>Months</th><th>Partners</th><th>Status</th><th><span className="sr-only">Action</span></th></tr></thead>
                <tbody>
                  {periods.map((p) => (
                    <tr key={p.periodId}>
                      <td className="num">{p.periodId}</td>
                      <td>{USAGE[p.usageType]?.label ?? p.usageType}</td>
                      <td>{monthYear(p.startDate)} to {monthYear(p.startDate, PERIOD_MONTHS - 1)}</td>
                      <td>{p.producerIds.map(nameOf).join(', ')}</td>
                      <td>{p.cancelledAt ? <Tag kind="blocked">Cancelled</Tag> : p.quorumComplete ? <Tag kind="valid">Complete</Tag> : <Tag kind="waiting">Answering</Tag>}</td>
                      <td>{!p.cancelledAt && <button className="link" onClick={() => handleCancelPeriod(p)}>Cancel</button>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>

        <section className="sec">
          <h2>Create epoch</h2>
          <p className="muted">Starts a new epoch of the usage type. It ends the open epoch of the same type; past epochs and their submissions are kept, and partners run a fresh key exchange.</p>
          <div className="row">
            <label className="field">Usage type
              <select value={epochUsage} onChange={(e) => setEpochUsage(e.target.value)}>
                {USAGE_TYPES.map((u) => <option key={u} value={u}>{USAGE[u].label}</option>)}
              </select>
            </label>
            <label className="field">Start month<input type="month" value={startMonth} onChange={(e) => setStartMonth(e.target.value)} /></label>
          </div>
          {epochPeriod
            ? <p>Reporting months {monthRange(`${startMonth}-01`, 4)}, in Quorum Check period {epochPeriod.periodId}. Partners: {epochPeriod.producerIds.map(nameOf).join(', ')}.</p>
            : <p className="why">No Quorum Check period of {USAGE[epochUsage].label} covers {monthRange(`${startMonth}-01`, 4)}. Open one first.</p>}
          {msg(epochMsg)}
          <div className="actions">
            {!confirmCreate
              ? <button className="btn btn-primary" onClick={() => setConfirmCreate(true)} disabled={!epochPeriod}>Create epoch</button>
              : <>
                  <span className="err">Create a new {USAGE[epochUsage].label} epoch? The open {USAGE[epochUsage].label} epoch ends.</span>
                  <button className="btn btn-primary" onClick={handleCreateEpoch} disabled={epochBusy}>{epochBusy ? 'Creating' : 'Confirm'}</button>
                  <button className="btn btn-secondary" onClick={() => setConfirmCreate(false)}>Cancel</button>
                </>}
          </div>
        </section>

        <section className="sec">
          <h2>Epochs</h2>
          {epochsError && <Alert kind="error" title="Epochs could not be loaded" lines={[`Server message: ${epochsError}`]} />}
          {epochs.length === 0 && !epochsLoading ? <p className="muted">No epochs yet.</p> : (
            <div className="atable-wrap">
              <table className="atable">
                <thead><tr><th>Epoch</th><th>Usage type</th><th>Month</th><th>Partners</th><th>Status</th></tr></thead>
                <tbody>
                  {epochs.map((e) => (
                    <tr key={e.epochId} className={`clickable${selectedEpochId === e.epochId ? ' selected' : ''}`} onClick={() => setSelectedEpochId(e.epochId)}>
                      <td className="num">{e.epochId}</td>
                      <td>{USAGE[e.usageType]?.label ?? e.usageType}</td>
                      <td>{monthYear(e.startDate)}</td>
                      <td>{e.producerCount}</td>
                      <td>{statusTag(e)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <div className="actions"><button className="btn btn-secondary" onClick={loadEpochs} disabled={epochsLoading}>{epochsLoading ? 'Refreshing' : 'Refresh'}</button></div>
        </section>

        {detailError && <Alert kind="error" title="The epoch could not be loaded" lines={[`Server message: ${detailError}`]} />}
        {detailLoading && <p className="muted">Loading the epoch.</p>}
        {epochDetail && !detailLoading && (
          <section className="sec">
            <div className="title-row" style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
              <h2>Epoch {epochDetail.epochId}: {USAGE[epochDetail.usageType]?.label}, {monthYear(epochDetail.startDate)}</h2>
              {statusTag(epochDetail)}
            </div>
            <p className="muted">{epochDetail.partnerCount} partners{epochDetail.quorumPeriod ? ` · Quorum Check period ${epochDetail.quorumPeriod.periodId}` : ' · No Quorum Check (created before periods)'}</p>

            {epochDetail.cancelledAt && (
              <Alert kind="warning" title={`Cancelled ${longDate(epochDetail.cancelledAt)}`} lines={[
                epochDetail.cancelReason && `Reason: ${epochDetail.cancelReason}`,
                'Partners in this epoch see it as cancelled.',
              ]}>
                {epochDetail.replacedByEpochId && <div className="actions"><button className="btn btn-secondary" onClick={() => setSelectedEpochId(epochDetail.replacedByEpochId)}>Open epoch {epochDetail.replacedByEpochId}</button></div>}
              </Alert>
            )}
            {msg(actionMsg)}

            {!epochDetail.isClosed && !epochDetail.replacedByEpochId && (!epochAction ? (
              <div className="actions">
                <button className="btn btn-secondary" onClick={() => setEpochAction('recreate')}>Recreate epoch</button>
                {!epochDetail.cancelledAt && <button className="btn btn-secondary" onClick={() => setEpochAction('cancel')}>Cancel epoch</button>}
              </div>
            ) : (
              <>
                <p className="err">
                  {epochAction === 'recreate'
                    ? `Cancel epoch ${epochDetail.epochId} and create a new epoch with the same ${epochDetail.partnerCount} partners and start month? Key exchange starts over.`
                    : `Cancel epoch ${epochDetail.epochId}? Partners can no longer exchange keys or submit for it.`}
                  {' '}Partners will see that it was cancelled.
                </p>
                <label className="field field-wide">Reason shown to partners (optional)
                  <textarea rows={2} maxLength={500} value={cancelReason} onChange={(ev) => setCancelReason(ev.target.value)} />
                </label>
                <div className="actions">
                  <button className="btn btn-primary" onClick={handleEpochAction} disabled={actionBusy}>{actionBusy ? 'Working' : 'Confirm'}</button>
                  <button className="btn btn-secondary" onClick={() => setEpochAction(null)} disabled={actionBusy}>Back</button>
                </div>
              </>
            ))}

            <QuorumSection detail={epochDetail} />

            {epochDetail.quorumComplete === false ? null : epochDetail.missingProducers?.length > 0 ? (
              <>
                <p><b>Has not submitted.</b> Aggregates are hidden until every partner has submitted every cell that met the quorum.</p>
                <div className="atable-wrap">
                  <table className="atable">
                    <thead><tr><th>Partner</th><th>Object ID</th><th>Missing cells</th></tr></thead>
                    <tbody>
                      {epochDetail.missingProducers.map((p) => (
                        <tr key={p.producerId}>
                          <td><b>{p.displayName}</b></td>
                          <td className="mono">{p.producerId}</td>
                          <td>{formatMissingCells(p.missingCells, ignoredCellCount(epochDetail))}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            ) : epochDetail.aggregates?.length > 0 ? (
              <div className="atable-wrap">
                <table className="atable">
                  <thead><tr><th>Country</th><th>Month</th><th>Indicator</th><th>Segment</th><th>Total</th><th>Submissions</th></tr></thead>
                  <tbody>
                    {epochDetail.aggregates.map((r) => (
                      <tr key={`${r.country}-${r.month}-${r.indicator}-${r.segment}`}>
                        <td><b>{r.country}</b></td><td>{r.month}</td><td>{r.indicator}</td><td>{r.segment}</td>
                        <td className="num">{formatInt(r.total)}</td><td>{r.submissionCount}/{r.expectedSubmissions}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : <p className="muted">No aggregates yet.</p>}
          </section>
        )}
      </main>
    </Page>
  )
}

const shortSig = (sig) => (sig ?? '').slice(0, 12)
const sigColor = (sig) => `hsl(${parseInt((sig ?? '0').slice(0, 6), 16) % 360}, 60%, 85%)`

// Metrics sharing a signature were backed by exactly the same set of partners.
function QuorumSection({ detail }) {
  const cells = detail.quorumCells ?? []
  if (!detail.quorumPeriod) return null

  const answered = (detail.quorumPartners ?? []).filter((p) => p.done).length
  const groups = new Set(cells.filter((c) => c.participantCount > 0).map((c) => c.signature))

  return (
    <>
      <p><b>Quorum Check, period {detail.quorumPeriod.periodId}:</b> {answered} of {detail.quorumPartners?.length ?? 0} partners answered.</p>
      {detail.quorumComplete === false ? (
        <div className="atable-wrap">
          <table className="atable">
            <thead><tr><th>Partner</th><th>Answered</th></tr></thead>
            <tbody>
              {detail.quorumPartners.map((p) => (
                <tr key={p.producerId}><td><b>{p.displayName}</b></td><td>{p.done ? 'Done' : `${p.answeredMetrics} metrics`}</td></tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <>
          <p className="muted">{groups.size} distinct participation {groups.size === 1 ? 'composition' : 'compositions'}. Same signature = same set of partners (a masked fingerprint; who participates stays hidden). Metrics under 3 participants are ignored.</p>
          <div className="atable-wrap">
            <table className="atable">
              <thead><tr><th>Country</th><th>Indicator</th><th>Participants</th><th>Signature</th><th>Status</th></tr></thead>
              <tbody>
                {cells.map((c) => (
                  <tr key={`${c.country}-${c.indicator}`}>
                    <td><b>{c.country}</b></td>
                    <td>{c.indicator}</td>
                    <td>{c.unverified ? 'Masks did not cancel' : `${c.participantCount}/${detail.partnerCount}`}</td>
                    <td>{c.signature ? <code className="sig" title={c.signature} style={{ background: sigColor(c.signature) }}>{shortSig(c.signature)}</code> : '—'}</td>
                    <td>{c.ignored ? <Tag kind="blocked">Ignored</Tag> : <Tag kind="valid">Active</Tag>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </>
  )
}
