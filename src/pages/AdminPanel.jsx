import { Fragment, useState, useEffect, useCallback } from 'react'
import { useMsal } from '@azure/msal-react'
import * as api from '../utils/api'
import { CELL_COUNT, formatInt } from '../utils/csvUpload'

// Must match EpochLifecycle.MinParticipants in the API
const MIN_PARTICIPANTS = 3

function formatEpochDate(iso) {
  if (!iso) return '—'
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return '—'
  return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric', timeZone: 'UTC' })
}

function formatMissingCells(cells) {
  const n = cells?.length ?? 0
  if (n === 0) return '—'
  const preview = cells
    .slice(0, 6)
    .map((c) => `${c.country} ${c.month} ${c.indicator}/${c.segment}`)
  const more = n > 6 ? ' …' : ''
  return `${n} of ${CELL_COUNT} missing (${preview.join(', ')}${more})`
}

export default function AdminPanel() {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  const myId = account?.localAccountId ?? ''

  // ── Registered partners ─────────────────────────────────────────────────────
  const [registered, setRegistered] = useState([])      // { producerId, displayName }
  const [regLoading, setRegLoading] = useState(false)
  const [regError, setRegError] = useState(null)

  // selected OIDs
  const [selected, setSelected] = useState(new Set())

  // ── Epoch creation ──────────────────────────────────────────────────────────
  const [startMonth, setStartMonth] = useState(() => {
    const d = new Date()
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
  })
  const [epochBusy, setEpochBusy] = useState(false)
  const [epochResult, setEpochResult] = useState(null)
  const [epochError, setEpochError] = useState(null)
  const [confirmReset, setConfirmReset] = useState(false)

  // ── Clear all ───────────────────────────────────────────────────────────────
  const [clearBusy, setClearBusy] = useState(false)
  const [clearConfirm, setClearConfirm] = useState(false)

  // ── Epochs ──────────────────────────────────────────────────────────────────
  const [epochs, setEpochs] = useState([])
  const [epochsLoading, setEpochsLoading] = useState(false)
  const [epochsError, setEpochsError] = useState(null)
  const [selectedEpochId, setSelectedEpochId] = useState(null)
  const [epochDetail, setEpochDetail] = useState(null)
  const [detailLoading, setDetailLoading] = useState(false)
  const [detailError, setDetailError] = useState(null)

  const loadRegistered = useCallback(async () => {
    setRegLoading(true)
    setRegError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const data = await api.getProducers(token)
      setRegistered(data.filter((p) => p.producerId !== myId))
    } catch (e) {
      setRegError(e.message)
    } finally {
      setRegLoading(false)
    }
  }, [instance, account, myId])

  const loadEpochs = useCallback(async () => {
    setEpochsLoading(true)
    setEpochsError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const data = await api.adminGetEpochs(token)
      const list = data.epochs ?? []
      setEpochs(list)
      setSelectedEpochId((prev) => {
        if (prev && list.some((e) => e.epochId === prev)) return prev
        return list[0]?.epochId ?? null
      })
    } catch (e) {
      setEpochsError(e.message)
    } finally {
      setEpochsLoading(false)
    }
  }, [instance, account])

  const loadEpochDetail = useCallback(async (epochId) => {
    if (!epochId) {
      setEpochDetail(null)
      return
    }
    setDetailLoading(true)
    setDetailError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const data = await api.adminGetEpochDetail(epochId, token)
      setEpochDetail(data)
    } catch (e) {
      setEpochDetail(null)
      setDetailError(e.message)
    } finally {
      setDetailLoading(false)
    }
  }, [instance, account])

  useEffect(() => {
    loadRegistered()
    loadEpochs()
  }, [])

  useEffect(() => {
    loadEpochDetail(selectedEpochId)
  }, [selectedEpochId, loadEpochDetail])

  const toggleSelect = (id) => setSelected((prev) => {
    const next = new Set(prev)
    next.has(id) ? next.delete(id) : next.add(id)
    return next
  })

  const toggleAll = () => {
    if (selected.size === registered.length) {
      setSelected(new Set())
    } else {
      setSelected(new Set(registered.map((p) => p.producerId)))
    }
  }

  const handleCreateEpoch = async () => {
    setEpochBusy(true)
    setEpochResult(null)
    setEpochError(null)
    try {
      const producers = [...selected].map((id) => {
        const p = registered.find((r) => r.producerId === id)
        return { producerId: id, displayName: p?.displayName ?? id }
      })
      const token = await api.acquireApiToken(instance, account)
      const data = await api.adminResetAndCreateEpoch({ startMonth, producers }, token)
      setEpochResult(data)
      setSelected(new Set())
      await loadRegistered()
      await loadEpochs()
      if (data.epoch?.epochId) setSelectedEpochId(data.epoch.epochId)
    } catch (e) {
      setEpochError(e.message)
    } finally {
      setEpochBusy(false)
      setConfirmReset(false)
    }
  }

  const handleClearAll = async () => {
    setClearBusy(true)
    try {
      const token = await api.acquireApiToken(instance, account)
      await api.adminReset(token)
      setRegistered([])
      setSelected(new Set())
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
  const selectedList = registered.filter((p) => selected.has(p.producerId))
  const canCreateEpoch = selected.size >= MIN_PARTICIPANTS
  const missingCount = MIN_PARTICIPANTS - selected.size

  return (
    <div className="animate-fade-in">
      <div className="page-header">
        <h1 className="page-title">⚙️ Admin Panel</h1>
        <p className="page-subtitle">Manage epochs, producers, and view aggregate results</p>
      </div>

      {/* Registered Partners */}
      <div className="card">
        <div className="card-header">
          <span className="card-icon">👥</span>
          <h2 className="card-title">Registered Partners</h2>
          <button className="btn btn-secondary" style={{ marginLeft: 'auto', padding: '0.35rem 0.75rem', fontSize: '0.8rem' }}
            onClick={loadRegistered} disabled={regLoading}>
            {regLoading ? '⏳' : '↻ Refresh'}
          </button>
        </div>

        <div className="info-box" style={{ marginBottom: '1rem' }}>
          <span className="info-box-icon">ℹ️</span>
          Partners appear here automatically when they log in and visit the Protocol page.
        </div>

        {regError && (
          <div className="info-box error" style={{ marginBottom: '0.75rem' }}>
            ⚠️ {regError}
          </div>
        )}

        {registered.length === 0 && !regLoading ? (
          <div className="text-muted" style={{ textAlign: 'center', padding: '1.5rem 0' }}>
            No partners registered yet. Partners must log in and visit the Protocol page first.
          </div>
        ) : (
          <table className="results-table" style={{ marginBottom: '1rem' }}>
            <thead>
              <tr>
                <th style={{ width: 40 }}>
                  <input type="checkbox" checked={allSelected} onChange={toggleAll} />
                </th>
                <th>Display Name</th>
                <th>OID (Producer ID)</th>
                <th>Joined</th>
              </tr>
            </thead>
            <tbody>
              {registered.map((p) => (
                <tr key={p.producerId} style={{ cursor: 'pointer' }} onClick={() => toggleSelect(p.producerId)}>
                  <td onClick={(e) => e.stopPropagation()}>
                    <input type="checkbox" checked={selected.has(p.producerId)}
                      onChange={() => toggleSelect(p.producerId)} />
                  </td>
                  <td style={{ fontWeight: 600 }}>{p.displayName}</td>
                  <td className="text-muted" style={{ fontFamily: 'monospace', fontSize: '0.8rem' }}>{p.producerId}</td>
                  <td className="text-muted" style={{ fontSize: '0.85rem' }}>
                    {p.joinedDate ? new Date(p.joinedDate).toLocaleDateString() : '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}

        {/* Clear All Data */}
        <div style={{ borderTop: '1px solid var(--gray-200)', paddingTop: '1rem', marginTop: '0.5rem' }}>
          {!clearConfirm ? (
            <button className="btn btn-secondary text-danger" style={{ fontSize: '0.8rem' }}
              onClick={() => setClearConfirm(true)}>
              🗑 Clear All Data
            </button>
          ) : (
            <div style={{ display: 'flex', gap: '0.75rem', alignItems: 'center' }}>
              <span className="text-danger" style={{ fontSize: '0.875rem' }}>Wipes all producers, keys, ciphertexts and submissions.</span>
              <button className="btn btn-primary" onClick={handleClearAll} disabled={clearBusy}>
                {clearBusy ? '⏳' : '⚠️ Confirm Clear'}
              </button>
              <button className="btn btn-secondary" onClick={() => setClearConfirm(false)}>Cancel</button>
            </div>
          )}
        </div>
      </div>

      {/* Create Epoch */}
      <div className="card">
        <div className="card-header">
          <span className="card-icon">🔄</span>
          <h2 className="card-title">Create Epoch</h2>
        </div>
        <div className="info-box warn" style={{ marginBottom: '1rem' }}>
          ⚠️ Starts a new epoch with the selected partners. Previous epochs and their submissions are kept.
          Keys and ciphertexts are cleared so partners re-run key exchange.
        </div>

        <div className="form-group">
          <label className="form-label">Start Month</label>
          <input type="month" className="form-input" style={{ maxWidth: 200 }}
            value={startMonth} onChange={(e) => setStartMonth(e.target.value)} />
        </div>

        {selected.size === 0 ? (
          <div className="text-muted" style={{ padding: '0.75rem 0' }}>
            Select at least {MIN_PARTICIPANTS} partners above to continue.
          </div>
        ) : (
          <div className="text-muted" style={{ marginBottom: '1rem', fontSize: '0.875rem' }}>
            {selected.size} partner{selected.size !== 1 ? 's' : ''} selected: {selectedList.map(p => p.displayName).join(', ')}
            {!canCreateEpoch && (
              <div className="text-danger" style={{ marginTop: '0.25rem' }}>
                Select {missingCount} more partner{missingCount !== 1 ? 's' : ''} — an epoch needs at least {MIN_PARTICIPANTS}.
              </div>
            )}
          </div>
        )}

        {epochError && (
          <div className="info-box error" style={{ marginBottom: '0.75rem' }}>
            ⚠️ {epochError}
          </div>
        )}

        {epochResult && (
          <div className="info-box ok" style={{ marginBottom: '0.75rem' }}>
            ✅ Epoch {epochResult.epoch?.epochId} created with {epochResult.producers?.length} producers.
          </div>
        )}

        {!confirmReset ? (
          <button className="btn btn-primary" onClick={() => setConfirmReset(true)} disabled={!canCreateEpoch}>
            🔄 Create Epoch
          </button>
        ) : (
          <div style={{ display: 'flex', gap: '0.75rem', alignItems: 'center' }}>
            <span className="text-danger">Create a new epoch? Keys and ciphertexts will be reset; past epochs stay.</span>
            <button className="btn btn-primary" onClick={handleCreateEpoch} disabled={epochBusy}>
              {epochBusy ? '⏳ Working…' : '⚠️ Confirm'}
            </button>
            <button className="btn btn-secondary" onClick={() => setConfirmReset(false)}>Cancel</button>
          </div>
        )}
      </div>

      {/* Epochs */}
      <div className="card">
        <div className="card-header">
          <span className="card-icon">📅</span>
          <h2 className="card-title">Epochs</h2>
          <button className="btn btn-secondary" style={{ marginLeft: 'auto', padding: '0.35rem 0.75rem', fontSize: '0.8rem' }}
            onClick={loadEpochs} disabled={epochsLoading}>
            {epochsLoading ? '⏳' : '↻ Refresh'}
          </button>
        </div>

        {epochsError && (
          <div className="info-box error">
            ⚠️ {epochsError}
          </div>
        )}

        {epochs.length === 0 && !epochsLoading ? (
          <div className="text-muted" style={{ textAlign: 'center', padding: '1rem' }}>No epochs yet.</div>
        ) : (
          <table className="results-table" style={{ marginBottom: epochDetail || detailLoading ? '1.5rem' : 0 }}>
            <thead>
              <tr>
                <th>Epoch</th>
                <th>Producers</th>
                <th>Date</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {epochs.map((e) => (
                <tr
                  key={e.epochId}
                  className={`clickable${selectedEpochId === e.epochId ? ' selected' : ''}`}
                  onClick={() => setSelectedEpochId(e.epochId)}
                >
                  <td style={{ fontWeight: 600 }}>{e.epochId}</td>
                  <td>{e.producerCount}</td>
                  <td>{formatEpochDate(e.startDate)}</td>
                  <td>
                    <span className={`status-badge ${e.isClosed ? 'pending' : 'success'}`}>
                      {e.isClosed ? 'Closed' : e.quorumComplete === false ? 'Quorum Check' : 'Open'}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}

        {detailError && (
          <div className="info-box error">
            ⚠️ {detailError}
          </div>
        )}

        {detailLoading && <div className="text-muted" style={{ padding: '0.5rem 0' }}>Loading epoch…</div>}

        {epochDetail && !detailLoading && (
          <>
            <div className="text-muted" style={{ fontSize: '0.875rem', marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: '0.5rem', flexWrap: 'wrap' }}>
              <span className="text-strong" style={{ fontWeight: 600 }}>Epoch {epochDetail.epochId}</span>
              {epochDetail.isClosed && <span className="status-badge pending">Closed</span>}
              <span>· {epochDetail.partnerCount} producers</span>
              {epochDetail.quorumComplete === false && <span className="status-badge warn">Quorum Check</span>}
            </div>

            <QuorumSection detail={epochDetail} />

            {epochDetail.quorumComplete === false ? null : epochDetail.missingProducers?.length > 0 ? (
              <>
                <div className="info-box warn">
                  Aggregates are hidden until every producer has submitted every cell.
                </div>
                <div className="text-strong" style={{ fontWeight: 600, marginBottom: '0.75rem' }}>Has not submitted</div>
                <table className="results-table">
                  <thead>
                    <tr>
                      <th>Producer</th>
                      <th>OID</th>
                      <th>Missing cells</th>
                    </tr>
                  </thead>
                  <tbody>
                    {epochDetail.missingProducers.map((p) => (
                      <tr key={p.producerId}>
                        <td style={{ fontWeight: 600, fontFamily: 'Inter, sans-serif' }}>{p.displayName}</td>
                        <td className="text-muted" style={{ fontSize: '0.8rem' }}>{p.producerId}</td>
                        <td style={{ fontSize: '0.8rem' }}>
                          {formatMissingCells(p.missingCells)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            ) : epochDetail.aggregates?.length > 0 ? (
              <table className="results-table">
                <thead>
                  <tr>
                    <th>Country</th>
                    <th>Month</th>
                    <th>Indicator</th>
                    <th>Segment</th>
                    <th>Total</th>
                    <th>Submissions</th>
                  </tr>
                </thead>
                <tbody>
                  {epochDetail.aggregates.map((r) => (
                    <tr key={`${r.country}-${r.month}-${r.indicator}-${r.segment}`}>
                      <td style={{ fontWeight: 600 }}>{r.country}</td>
                      <td>{r.month}</td>
                      <td>{r.indicator}</td>
                      <td>{r.segment}</td>
                      <td className="text-success">{formatInt(r.total)}</td>
                      <td>{r.submissionCount}/{r.expectedSubmissions}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <div className="text-muted" style={{ textAlign: 'center', padding: '1rem' }}>No producers in this epoch.</div>
            )}
          </>
        )}
      </div>
    </div>
  )
}

const shortSig = (sig) => (sig ?? '').slice(0, 12)
const sigColor = (sig) => `hsl(${parseInt((sig ?? '0').slice(0, 6), 16) % 360}, 60%, 85%)`

// Metrics sharing a signature were backed by exactly the same set of partners.
function QuorumSection({ detail }) {
  const cells = detail.quorumCells ?? []
  if (cells.length === 0 && !(detail.quorumPartners?.length > 0)) return null

  const answered = (detail.quorumPartners ?? []).filter((p) => p.done).length
  const groups = new Map()
  for (const c of cells) {
    if (c.participantCount === 0) continue
    groups.set(c.signature, (groups.get(c.signature) ?? 0) + 1)
  }

  return (
    <div style={{ marginBottom: '1.5rem' }}>
      <div className="text-strong" style={{ fontWeight: 600, marginBottom: '0.5rem' }}>
        Quorum Check · {answered}/{detail.quorumPartners?.length ?? 0} partners answered
      </div>
      {detail.quorumComplete === false ? (
        <table className="results-table">
          <thead><tr><th>Producer</th><th>Answered cells</th></tr></thead>
          <tbody>
            {detail.quorumPartners.map((p) => (
              <tr key={p.producerId}>
                <td style={{ fontWeight: 600 }}>{p.displayName}</td>
                <td>{p.done ? 'Done' : `${p.answeredCells}`}</td>
              </tr>
            ))}
          </tbody>
        </table>
      ) : (
        <>
          <div className="text-muted" style={{ fontSize: '0.8rem', marginBottom: '0.5rem' }}>
            {groups.size} distinct participation composition(s). Same signature = same set of partners (a masked fingerprint;
            who participates stays hidden). Metrics under 3 participants are ignored.
          </div>
          <table className="results-table compact">
            <thead>
              <tr>
                <th>Country</th><th>Month</th><th>Indicator</th><th>Segment</th>
                <th>Participants</th><th>Signature</th><th>Status</th>
              </tr>
            </thead>
            <tbody>
              {cells.map((c) => {
                const key = `${c.country}-${c.month}-${c.indicator}-${c.segment}`
                return (
                  <Fragment key={key}>
                    <tr>
                      <td style={{ fontWeight: 600 }}>{c.country}</td>
                      <td>{c.month}</td>
                      <td>{c.indicator}</td>
                      <td>{c.segment}</td>
                      <td>{c.participantCount}/{detail.partnerCount}</td>
                      <td>
                        <code title={c.signature} style={{ background: sigColor(c.signature), padding: '2px 6px', borderRadius: 4 }}>
                          {shortSig(c.signature)}
                        </code>
                      </td>
                      <td>
                        <span className={`status-badge ${c.ignored ? 'error' : 'success'}`}>
                          {c.ignored ? 'Ignored' : 'Active'}
                        </span>
                      </td>
                    </tr>
                  </Fragment>
                )
              })}
            </tbody>
          </table>
        </>
      )}
    </div>
  )
}
