import { useState, useEffect, useCallback } from 'react'
import { useMsal } from '@azure/msal-react'
import { Page } from '../components/Shell'
import { Alert, Checkbox, Disclosure, Tag } from '../components/ui'
import * as api from '../utils/api'
import { QUARTER_MONTHS, SEGMENTS_BY_INDICATOR, cellCount, formatInt, monthCount } from '../utils/csvUpload'
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

// First day of the period's quarter `slot` (0-based)
const slotStart = (period, slot) => {
  const t = monthIndex(period.startDate) + slot * QUARTER_MONTHS
  return `${Math.floor(t / 12)}-${String((t % 12) + 1).padStart(2, '0')}-01`
}

// The quarter of the period an epoch reports in (0-based)
const slotOf = (period, epoch) => Math.floor((monthIndex(epoch.startDate) - monthIndex(period.startDate)) / QUARTER_MONTHS)

const blockedText = (p) => ({
  'period-cancelled': 'This period was cancelled.',
  'quorum-incomplete': 'Epochs can be created once every partner has answered the Quorum Check.',
  'epoch-open': `Epoch ${p.openEpochId} is still open. It closes when every partner has submitted; cancel it to move on.`,
  'period-full': `All ${p.epochCount} quarters have an epoch.`,
})[p.blockedReason] ?? null

function formatMissingCells(cells, required) {
  const n = cells?.length ?? 0
  if (n === 0) return '—'
  const preview = cells.slice(0, 6).map((c) => `${c.country} ${c.month} ${c.indicator}/${c.segment}`)
  return `${n} of ${required} missing (${preview.join(', ')}${n > 6 ? ' …' : ''})`
}

// Cells of the metrics that fell below the quorum: every segment of the indicator, in every month
const ignoredCellCount = (detail) => (detail.quorumCells ?? [])
  .filter((c) => c.ignored)
  .reduce((n, c) => n + monthCount(detail) * (SEGMENTS_BY_INDICATOR[c.indicator]?.length ?? 0), 0)

// Below-quorum cells get no data, so they are not required
const requiredCellCount = (detail) => cellCount(monthCount(detail)) - ignoredCellCount(detail)

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
  const [epochBusy, setEpochBusy] = useState(false)
  const [epochMsg, setEpochMsg] = useState(null)
  const [confirmPeriodId, setConfirmPeriodId] = useState(null) // period whose next epoch is being confirmed

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

  // Only one Quorum Check can be unfinished at a time; follow it until partners finish answering
  const activePeriod = periods.find((p) => !p.cancelledAt && !p.quorumComplete) ?? null
  useEffect(() => {
    if (!activePeriod) return
    const id = setInterval(loadPeriods, 10000)
    return () => clearInterval(id)
  }, [activePeriod?.periodId, loadPeriods])

  // Periods that still have work: answering, a quarter left to create, or an epoch still open
  const epochsOf = (p) => epochs.filter((e) => e.quorumPeriodId === p.periodId)
  const workingPeriods = periods
    .filter((p) => !p.cancelledAt && (p.nextSlot != null || epochsOf(p).some((e) => !e.isClosed && !e.cancelledAt)))
    .sort((a, b) => b.periodId - a.periodId)

  const refreshAll = () => Promise.all([loadPeriods(), loadEpochs()])

  const handleCreateEpoch = async (p) => {
    setEpochBusy(true)
    setEpochMsg(null)
    try {
      const data = await api.adminCreateNextEpoch(p.periodId, await token())
      setEpochMsg({ kind: 'ok', text: `Epoch ${data.epoch?.epochId} created for Q${p.nextSlot + 1}, ${monthRange(data.epoch?.startDate, QUARTER_MONTHS)}, with ${data.producers?.length} partners.` })
      await refreshAll()
      if (data.epoch?.epochId) setSelectedEpochId(data.epoch.epochId)
    } catch (e) {
      setEpochMsg({ kind: 'error', text: e.message })
      await loadPeriods()
    } finally {
      setEpochBusy(false)
      setConfirmPeriodId(null)
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
        await refreshAll()
        if (data.epoch?.epochId) setSelectedEpochId(data.epoch.epochId)
      } else {
        await api.adminCancelEpoch(epochId, reason, await token())
        setActionMsg({ kind: 'ok', text: `Epoch ${epochId} was cancelled.` })
        await refreshAll()
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
          <p className="meta">Open a Quorum Check period for a year, then create its quarterly epochs in order once every partner has answered.</p>
        </div>

        <section className="sec">
          <h2>Step 1 · Quorum Check period</h2>
          <p className="muted">A period covers {PERIOD_MONTHS} months of one usage type. Its partners answer the Quorum Check once, and every epoch in it has the same partners. Only one Quorum Check can be answering at a time.</p>
          {msg(periodMsg)}
          {activePeriod ? (
            <>
              <div className="title-row" style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
                <h3>Period {activePeriod.periodId}: {USAGE[activePeriod.usageType]?.label}, {monthYear(activePeriod.startDate)} to {monthYear(activePeriod.startDate, PERIOD_MONTHS - 1)}</h3>
                <Tag kind="waiting">Answering</Tag>
              </div>
              <p>Partners: {activePeriod.producerIds.map(nameOf).join(', ')}.</p>
              <p className="why">Partners are answering the Quorum Check. A new period can be opened once this one is complete or cancelled.</p>
              <div className="actions"><button className="btn btn-secondary" onClick={() => handleCancelPeriod(activePeriod)}>Cancel period</button></div>
            </>
          ) : (
            <>
              <p>Select at least {MIN_PARTICIPANTS} partners. Partners appear here automatically when they sign in.</p>
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
                <button className="btn btn-secondary" onClick={loadRegistered} disabled={regLoading}>{regLoading ? 'Refreshing' : 'Refresh partners'}</button>
              </div>
              {selected.size < MIN_PARTICIPANTS && <p className="why">Select {missingCount} more {missingCount === 1 ? 'partner' : 'partners'} above.</p>}
            </>
          )}
        </section>

        <section className="sec">
          <h2>Step 2 · Quarterly epochs</h2>
          <p className="muted">Each period has {PERIOD_MONTHS / QUARTER_MONTHS} quarterly epochs, created in order. The next one can be created when the Quorum Check is complete and no other epoch is open. An epoch closes by itself when every partner has submitted.</p>
          {msg(epochMsg)}
          {workingPeriods.length === 0 && <p className="muted">No period has epochs left to create. Open a Quorum Check period in step 1.</p>}
          {workingPeriods.map((p) => {
            const own = epochsOf(p)
            const quarters = Array.from({ length: p.epochCount ?? PERIOD_MONTHS / QUARTER_MONTHS }, (_, i) => {
              const inSlot = own.filter((e) => slotOf(p, e) === i)
              return { i, epoch: inSlot.find((e) => !e.cancelledAt) ?? inSlot[0] ?? null }
            })
            return (
              <div key={p.periodId} className="sec">
                <div className="title-row" style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
                  <h3>Period {p.periodId}: {USAGE[p.usageType]?.label}, {monthYear(p.startDate)} to {monthYear(p.startDate, PERIOD_MONTHS - 1)}</h3>
                  {p.quorumComplete ? <Tag kind="valid">Quorum Check complete</Tag> : <Tag kind="waiting">Answering</Tag>}
                </div>
                <div className="atable-wrap">
                  <table className="atable">
                    <thead><tr><th>Quarter</th><th>Months</th><th>Epoch</th><th>Status</th></tr></thead>
                    <tbody>
                      {quarters.map(({ i, epoch }) => (
                        <tr key={i} className={epoch ? `clickable${selectedEpochId === epoch.epochId ? ' selected' : ''}` : undefined}
                          onClick={epoch ? () => setSelectedEpochId(epoch.epochId) : undefined}>
                          <td><b>Q{i + 1}</b></td>
                          <td>{monthRange(epoch?.startDate ?? slotStart(p, i), epoch ? monthCount(epoch) : QUARTER_MONTHS)}</td>
                          <td className="num">{epoch?.epochId ?? '—'}</td>
                          <td>{epoch ? statusTag(epoch) : i === p.nextSlot ? <Tag kind="pending">Next</Tag> : <Tag kind="pending">Not created</Tag>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {p.nextSlot != null && (
                  <div className="actions">
                    {confirmPeriodId !== p.periodId
                      ? <button className="btn btn-primary" onClick={() => setConfirmPeriodId(p.periodId)} disabled={!p.canCreateEpoch}>
                          Create Q{p.nextSlot + 1} epoch ({monthRange(p.nextSlotStart, QUARTER_MONTHS)})
                        </button>
                      : <>
                          <span className="err">Create the Q{p.nextSlot + 1} epoch with the period's {p.producerCount} partners? They start a fresh key exchange.</span>
                          <button className="btn btn-primary" onClick={() => handleCreateEpoch(p)} disabled={epochBusy}>{epochBusy ? 'Creating' : 'Confirm'}</button>
                          <button className="btn btn-secondary" onClick={() => setConfirmPeriodId(null)}>Cancel</button>
                        </>}
                  </div>
                )}
                {!p.canCreateEpoch && blockedText(p) && <p className="why">{blockedText(p)}</p>}
              </div>
            )
          })}
          <div className="actions"><button className="btn btn-secondary" onClick={refreshAll} disabled={epochsLoading}>{epochsLoading ? 'Refreshing' : 'Refresh'}</button></div>
        </section>

        {detailError && <Alert kind="error" title="The epoch could not be loaded" lines={[`Server message: ${detailError}`]} />}
        {detailLoading && <p className="muted">Loading the epoch.</p>}
        {epochDetail && !detailLoading && (
          <section className="sec">
            <div className="title-row" style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
              <h2>Epoch {epochDetail.epochId}: {USAGE[epochDetail.usageType]?.label}, {monthRange(epochDetail.startDate, monthCount(epochDetail))}</h2>
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
                          <td>{formatMissingCells(p.missingCells, requiredCellCount(epochDetail))}</td>
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

        <section className="sec">
          <h2>History</h2>
          {epochsError && <Alert kind="error" title="Epochs could not be loaded" lines={[`Server message: ${epochsError}`]} />}
          <Disclosure id="history" label="Show all periods and epochs">
            <h3>Quorum Check periods</h3>
            {periods.length === 0 ? <p className="muted">No periods yet.</p> : (
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
            <h3>Epochs</h3>
            {epochs.length === 0 && !epochsLoading ? <p className="muted">No epochs yet.</p> : (
              <div className="atable-wrap">
                <table className="atable">
                  <thead><tr><th>Epoch</th><th>Usage type</th><th>Months</th><th>Period</th><th>Partners</th><th>Status</th></tr></thead>
                  <tbody>
                    {epochs.map((e) => (
                      <tr key={e.epochId} className={`clickable${selectedEpochId === e.epochId ? ' selected' : ''}`} onClick={() => setSelectedEpochId(e.epochId)}>
                        <td className="num">{e.epochId}</td>
                        <td>{USAGE[e.usageType]?.label ?? e.usageType}</td>
                        <td>{monthRange(e.startDate, monthCount(e))}</td>
                        <td>{e.quorumPeriodId ?? '—'}</td>
                        <td>{e.producerCount}</td>
                        <td>{statusTag(e)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </Disclosure>
        </section>

        <section className="sec">
          <h2>Danger zone</h2>
          <div className="actions">
            {!clearConfirm
              ? <button className="btn btn-secondary" onClick={() => setClearConfirm(true)}>Clear all data</button>
              : <>
                  <span className="err">Wipes all partners, periods, epochs, keys, ciphertexts, answers and submissions.</span>
                  <button className="btn btn-primary" onClick={handleClearAll} disabled={clearBusy}>{clearBusy ? 'Clearing' : 'Confirm clear'}</button>
                  <button className="btn btn-secondary" onClick={() => setClearConfirm(false)}>Cancel</button>
                </>}
          </div>
        </section>
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
