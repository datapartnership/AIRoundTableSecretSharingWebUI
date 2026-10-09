import { useCallback, useEffect, useRef, useState } from 'react'
import { useSearchParams } from 'react-router-dom'
import { useMsal } from '@azure/msal-react'
import * as api from '../utils/api'
import { getDeviceId, loadLocalKeyPair, keyStatus, KEY_STATUS } from '../utils/localCrypto'
import EpochProtocol from './EpochProtocol'

const POLL_MS = 2000

export default function EpochDashboard({ isAdmin = false }) {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  const myId = account?.localAccountId ?? ''
  const deviceId = getDeviceId()

  const [epochs, setEpochs] = useState(null)
  const [epochError, setEpochError] = useState(null)
  const [statuses, setStatuses] = useState({})
  const [searchParams, setSearchParams] = useSearchParams()
  const selectedId = Number(searchParams.get('epoch')) || null
  const lastSelectedRef = useRef(null)

  const refresh = useCallback(async () => {
    try {
      const token = await api.acquireApiToken(instance, account)
      const data = await api.getEpochs(token)
      const list = (data.epochs ?? []).map((e) => (isAdmin ? { ...e, isEligible: false } : e))
      setEpochs(list)
      setEpochError(null)

      const eligible = list.filter((e) => e.isEligible && !e.cancelledAt)
      const results = await Promise.all(eligible.map((e) =>
        api.getKeyExchangeStatus(e.epochId, deviceId, token)
          .then((s) => [e.epochId, s])
          .catch((err) => [e.epochId, { error: err.message }])
      ))
      setStatuses(Object.fromEntries(results))
    } catch (e) {
      console.error('[dashboard] refresh failed', e)
      setEpochError(e.message)
    }
  }, [instance, account?.homeAccountId, deviceId, isAdmin])

  useEffect(() => {
    if (!account) return
    refresh()
    const id = setInterval(refresh, POLL_MS)
    return () => clearInterval(id)
  }, [account?.homeAccountId, refresh])

  const onKeyStatusChange = useCallback((epochId, s) => {
    setStatuses((prev) => ({ ...prev, [epochId]: s }))
  }, [])

  const select = useCallback((epochId) => {
    setSearchParams(epochId ? { epoch: String(epochId) } : {}, { replace: true })
  }, [setSearchParams])

  const eligible = (epochs ?? []).filter((e) => e.isEligible && !e.cancelledAt)
  const selected = (epochs ?? []).find((e) => e.epochId === selectedId) ?? null
  if (selected) lastSelectedRef.current = selected

  // Auto-select when there's nothing (still open) selected and exactly one epoch to act on
  useEffect(() => {
    if (!epochs || selected || eligible.length !== 1) return
    select(eligible[0].epochId)
  }, [epochs, selected, eligible.length, eligible[0]?.epochId, select])

  const closedSelection = !selected && selectedId && lastSelectedRef.current?.epochId === selectedId
    ? { ...lastSelectedRef.current, isClosed: true }
    : null
  const activeEpoch = selected ?? closedSelection

  const badgeFor = (epoch) => {
    if (epoch.cancelledAt) return <span className="status-badge error">Cancelled by admin</span>
    if (isAdmin) return <span className="status-badge pending">View only (admin)</span>
    if (!epoch.isEligible) return <span className="status-badge pending">Not a participant</span>
    const s = statuses[epoch.epochId]
    if (s?.error) return <span className="status-badge error" title={s.error}>Key status unavailable</span>
    const state = keyStatus(loadLocalKeyPair(myId, epoch.epochId, deviceId), s ?? null, myId)
    return <span className={`status-badge ${KEY_STATUS[state].tone}`}>{KEY_STATUS[state].label}</span>
  }

  return (
    <div className="animate-fade-in">
      <div className="page-header">
        <h1 className="page-title">Epochs</h1>
        <p className="page-subtitle">Pick an open epoch to exchange keys and submit privacy-preserving metrics</p>
      </div>

      <div className="card" style={{ marginBottom: '1.5rem' }}>
        <div className="card-header">
          <span className="card-icon">🗓️</span>
          <h2 className="card-title">Open epochs</h2>
        </div>
        {isAdmin && (
          <div className="info-box" style={{ marginBottom: '1rem' }}>
            <span className="info-box-icon">ℹ️</span>
            Admins cannot participate in epochs. Use the Admin Panel to create epochs and follow submissions.
          </div>
        )}
        {epochError && <div className="info-box error">⚠️ {epochError}</div>}
        {epochs === null && !epochError && <div className="text-muted">Loading epochs…</div>}
        {epochs?.length === 0 && (
          <div className="text-muted">No epochs are currently open. This page updates when an admin creates one.</div>
        )}
        {(epochs ?? []).map((epoch) => {
          const s = statuses[epoch.epochId]
          const isSelected = epoch.epochId === activeEpoch?.epochId
          return (
            <div
              key={epoch.epochId}
              className="calc-step"
              onClick={epoch.isEligible ? () => select(epoch.epochId) : undefined}
              style={{
                cursor: epoch.isEligible ? 'pointer' : 'default',
                opacity: epoch.isEligible ? 1 : 0.7,
                background: isSelected ? 'rgba(86, 195, 196, 0.08)' : undefined,
                borderRadius: 8,
                paddingLeft: '0.75rem',
                paddingRight: '0.75rem',
              }}
            >
              <div className="step-number">{epoch.epochId}</div>
              <div className="step-content" style={{ flex: 1 }}>
                <div className="step-title" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', flexWrap: 'wrap' }}>
                  {epoch.startDate?.slice(0, 10) ?? 'Unknown start'}
                  {badgeFor(epoch)}
                  {isSelected && <span className="status-badge success">Selected</span>}
                </div>
                <div className="step-description">
                  {epoch.cancelledAt
                    ? `Cancelled ${formatDate(epoch.cancelledAt)}${epoch.replacedByEpochId ? ` · replaced by epoch ${epoch.replacedByEpochId}` : ''}`
                    : epoch.isEligible && s && !s.error
                    ? `${s.registeredCount}/${s.expectedCount} partners registered keys`
                    : `${epoch.producerCount ?? epoch.producerIds?.length ?? 0} partners`}
                </div>
              </div>
            </div>
          )
        })}
      </div>

      {activeEpoch?.cancelledAt && (
        <CancelledEpochNotice epoch={activeEpoch} onSelect={select} />
      )}

      {activeEpoch && !activeEpoch.cancelledAt && (
        <EpochProtocol
          key={activeEpoch.epochId}
          epoch={activeEpoch}
          onRefresh={refresh}
          onKeyStatusChange={onKeyStatusChange}
          isAdmin={isAdmin}
        />
      )}
    </div>
  )
}

function formatDate(iso) {
  const d = new Date(iso)
  return Number.isNaN(d.getTime()) ? String(iso) : d.toLocaleString()
}

function CancelledEpochNotice({ epoch, onSelect }) {
  return (
    <div className="card">
      <div className="card-header">
        <span className="card-icon">⛔</span>
        <h2 className="card-title">Epoch {epoch.epochId} was cancelled</h2>
      </div>
      <div className="info-box warn">
        <span className="info-box-icon">⚠️</span>
        <div>
          An admin cancelled this epoch on {formatDate(epoch.cancelledAt)}. No further key exchange, Quorum Check answers
          or submissions are accepted for it.
          {epoch.cancelReason && <div style={{ marginTop: '0.5rem' }}><strong>Reason:</strong> {epoch.cancelReason}</div>}
          {epoch.replacedByEpochId && (
            <div style={{ marginTop: '0.5rem' }}>
              It was replaced by{' '}
              <button type="button" className="btn btn-secondary" onClick={() => onSelect(epoch.replacedByEpochId)}>
                Epoch {epoch.replacedByEpochId}
              </button>
              {' '}— continue there.
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
