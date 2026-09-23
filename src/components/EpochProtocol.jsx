import { Fragment, useState, useEffect, useCallback, useRef } from 'react'
import { useMsal } from '@azure/msal-react'
import * as api from '../utils/api'
import { generateMlKemKeyPair, encapsulate, decapsulate } from '../utils/crypto'
import { calculateMaskedValue } from '../utils/noise'
import {
  parseAndValidateCsv,
  isCsvFile,
  epochMonths,
  emptyCsvResult,
  rewriteSampleMonths,
  CELL_COUNT,
  COUNTRIES,
  SERIES,
  HEADERS,
  ERROR_DISPLAY_CAP,
  formatInt,
} from '../utils/csvUpload'
import {
  getDeviceId, loadLocalCrypto, wipeLocalCrypto, persistKeyPair, persistSecrets, persistSent, persistSeen,
  keyStatus, keyFingerprint, KEY_STATUS,
} from '../utils/localCrypto'

const POLL_MS = 2000
const TAG = '[protocol]'

function log(event, data) {
  if (data === undefined) console.log(TAG, event)
  else console.log(TAG, event, data)
}

function warn(event, data) {
  if (data === undefined) console.warn(TAG, event)
  else console.warn(TAG, event, data)
}

function fail(event, err, data) {
  if (data === undefined) console.error(TAG, event, err)
  else console.error(TAG, event, err, data)
}

const KEY_RECOVERY_BLOCKED_MESSAGE = {
  mismatch: 'This browser’s key does not match the key registered on the server, and it can’t be replaced automatically because the key exchange has already started. Ask an admin to recreate the epoch.',
  'server-only': 'This browser has no private key for the key registered on the server, and it can’t be replaced automatically because the key exchange has already started. Ask an admin to recreate the epoch.',
  'other-device': 'Your key for this epoch was registered from another browser or device, and it can’t be replaced automatically because the key exchange has already started. Continue there, or ask an admin to recreate the epoch.',
}

const ROTATE_STATES = new Set(['mismatch', 'server-only', 'other-device'])

// Mount with key={epoch.epochId} so all state resets when the selected epoch changes.
export default function EpochProtocol({ epoch, onRefresh, onKeyStatusChange, isAdmin = false }) {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  // Azure AD OID matches the JWT sub claim the API uses for metrics
  const myId = account?.localAccountId ?? ''
  const deviceId = getDeviceId()
  const epochId = epoch.epochId

  const [step, setStep] = useState(1)
  const [hydrated, setHydrated] = useState(false)

  // Key pair state
  const [keyPair, setKeyPair] = useState(null)
  const [keyBusy, setKeyBusy] = useState(false)
  const [keyError, setKeyError] = useState(null)
  const [keyRotating, setKeyRotating] = useState(false)
  const [recoveryBlocked, setRecoveryBlocked] = useState(false)
  const [fingerprints, setFingerprints] = useState({ local: null, server: null })

  // Polling
  const [status, setStatus] = useState(null)
  const [partnerKeys, setPartnerKeys] = useState([])
  const [pollError, setPollError] = useState(null)

  // Ciphertext exchange
  const [sentTo, setSentTo] = useState(new Set())       // partners I encapsulated for
  const [sharedSecrets, setSharedSecrets] = useState(new Map())
  const [ctSeen, setCtSeen] = useState(new Map())       // senderId → ciphertext blob we already decapped
  const [encapBusy, setEncapBusy] = useState(false)
  const [encapError, setEncapError] = useState(null)

  // CSV upload
  const [submittedCells, setSubmittedCells] = useState(new Set())
  const [csvFileName, setCsvFileName] = useState('')
  const [csvResult, setCsvResult] = useState(null)
  const [submitBusy, setSubmitBusy] = useState(false)
  const [submitError, setSubmitError] = useState(null)
  const [loadError, setLoadError] = useState(null)

  const keyPairRef = useRef(null)
  const secretsRef = useRef(new Map())
  const sentToRef = useRef(new Set())
  const ctSeenRef = useRef(new Map())
  const onKeyStatusChangeRef = useRef(onKeyStatusChange)
  // Poll responses requested before the last key change describe the old key and must not trigger recovery
  const keyChangedAtRef = useRef(0)
  const lastRecoveryStatusRef = useRef(null)
  const lastDecapStatusRef = useRef(null)
  keyPairRef.current = keyPair
  secretsRef.current = sharedSecrets
  sentToRef.current = sentTo
  ctSeenRef.current = ctSeen
  onKeyStatusChangeRef.current = onKeyStatusChange

  const applyClearedCrypto = () => {
    warn('local crypto wiped')
    setKeyPair(null)
    setSharedSecrets(new Map())
    setSentTo(new Set())
    setCtSeen(new Map())
    setStep(1)
    setEncapError(null)
    setKeyError(null)
  }

  // ── Restore persisted state before any auto-run ──────────────────────────────
  useEffect(() => {
    if (!myId) return
    const local = loadLocalCrypto(myId, epochId, deviceId)
    log('restored local crypto', {
      epochId,
      hasKeyPair: !!local.keyPair,
      secretPartners: [...local.secrets.keys()],
      sentTo: [...local.sentTo],
      seenFrom: [...local.ctSeen.keys()],
    })
    setKeyPair(local.keyPair)
    setSharedSecrets(local.secrets)
    setSentTo(local.sentTo)
    setCtSeen(local.ctSeen)
    setHydrated(true)
  }, [myId, epochId, deviceId])

  // ── Poll key exchange while setup is in progress ─────────────────────────────
  const epochInactive = isAdmin || !!epoch.isClosed || epoch.isEligible === false

  useEffect(() => {
    if (!hydrated || !myId || epochInactive) return
    let alive = true

    const poll = async () => {
      try {
        const token = await api.acquireApiToken(instance, account)
        if (step > 3) return

        const requestedAt = Date.now()
        const [rawStatus, pk, sent] = await Promise.all([
          api.getKeyExchangeStatus(epochId, deviceId, token),
          api.getPartnerKeys(epochId, deviceId, token),
          api.getSentCiphertexts(epochId, deviceId, token).catch(() => ({ ciphertexts: [] })),
        ])
        if (!alive) return

        const s = { ...rawStatus, requestedAt }
        setPollError(null)
        setStatus(s)
        setPartnerKeys(pk.partnerKeys ?? [])
        onKeyStatusChangeRef.current?.(epochId, s)

        const serverSent = (sent.ciphertexts ?? []).map((c) => c.recipientId)
        if (serverSent.length) {
          setSentTo((prev) => {
            const next = new Set(prev)
            for (const id of serverSent) next.add(id)
            persistSent(myId, epochId, deviceId, next)
            return next
          })
        }

        log('poll', {
          step,
          epochId,
          producers: epoch.producerIds,
          registered: `${s.registeredCount}/${s.expectedCount}`,
          missing: s.missingPartners,
          partnerKeys: (pk.partnerKeys ?? []).map((p) => p.producerId),
          ciphertexts: `${s.actualCiphertexts}/${s.expectedCiphertexts}`,
          exchangeComplete: s.isCiphertextExchangeComplete,
          serverSentTo: serverSent,
          localSentTo: [...sentToRef.current],
          secretPartners: [...secretsRef.current.keys()],
          hasLocalKey: !!keyPairRef.current,
          hasServerKey: !!s.myPublicKeyBase64,
        })
      } catch (e) {
        fail('poll failed', e)
        if (alive) setPollError(`Can’t reach the server to check key exchange progress: ${e.message}`)
      }
    }

    poll()
    const id = setInterval(poll, POLL_MS)
    return () => { alive = false; clearInterval(id) }
  }, [hydrated, step, myId, epochId, epochInactive, instance, account?.homeAccountId, deviceId])

  // ── Key fingerprints for the status bar ──────────────────────────────────────
  useEffect(() => {
    let alive = true
    Promise.all([keyFingerprint(keyPair?.ekBase64), keyFingerprint(status?.myPublicKeyBase64)])
      .then(([local, server]) => { if (alive) setFingerprints({ local, server }) })
    return () => { alive = false }
  }, [keyPair?.ekBase64, status?.myPublicKeyBase64])

  // ── Step 1: generate & register key pair ─────────────────────────────────────
  const generateAndRegister = useCallback(async () => {
    setKeyBusy(true)
    setKeyError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const existing = keyPairRef.current
      const kp = existing ?? await generateMlKemKeyPair()
      log(existing ? 're-registering existing key' : 'generated new key pair', { myId, epochId })
      await api.registerPublicKey(epochId, deviceId, kp.ekBase64, token)
      log('public key registered')
      keyChangedAtRef.current = Date.now()
      if (!existing) {
        persistKeyPair(myId, epochId, deviceId, kp)
        setKeyPair(kp)
      }
    } catch (e) {
      fail('key generate/register failed', e)
      setKeyError(e.message)
    } finally {
      setKeyBusy(false)
    }
  }, [instance, account, myId, epochId, deviceId])

  // ── Step 1 recovery: replace a missing/mismatched key while nothing depends on it ─
  const rotateKey = useCallback(async () => {
    setKeyBusy(true)
    setKeyRotating(true)
    setKeyError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const kp = await generateMlKemKeyPair()
      warn('rotating key', { myId, epochId })
      await api.rotatePublicKey(epochId, deviceId, kp.ekBase64, token)
      keyChangedAtRef.current = Date.now()
      wipeLocalCrypto(myId, epochId, deviceId)
      applyClearedCrypto()
      persistKeyPair(myId, epochId, deviceId, kp)
      setKeyPair(kp)
      log('key rotated')
    } catch (e) {
      if (e.status === 409 && e.code === 'exchange-started') {
        warn('key rotation blocked — exchange already started', { myId, epochId })
        setRecoveryBlocked(true)
      } else {
        fail('key rotation failed', e)
        setKeyError(e.message)
      }
    } finally {
      setKeyBusy(false)
      setKeyRotating(false)
    }
  }, [instance, account, myId, epochId, deviceId])

  // ── Step 2: encapsulate for all smaller-ID partners (never overwrite) ───────
  const performEncapsulation = useCallback(async () => {
    setEncapBusy(true)
    setEncapError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const targets = partnerKeys.filter((pk) => pk.producerId < myId)
      log('encapsulation start', { targets: targets.map((p) => p.producerId) })
      const newSecrets = new Map(secretsRef.current)
      const newSent = new Set(sentToRef.current)

      for (const pk of targets) {
        if (newSecrets.has(pk.producerId) && newSent.has(pk.producerId)) {
          log('encapsulation skip (already done)', { recipient: pk.producerId })
          continue
        }

        if (newSent.has(pk.producerId) && !newSecrets.has(pk.producerId)) {
          throw new Error(
            `A ciphertext for ${pk.producerId} is already on the server, but this browser lost the shared secret. Recreate the epoch so both sides start over.`
          )
        }

        const { ctBase64, sharedSecret } = await encapsulate(pk.publicKeyBase64)
        await api.postCiphertext(epochId, deviceId, pk.producerId, pk.deviceId, ctBase64, token)
        log('encapsulated + posted ciphertext', { recipient: pk.producerId, ctBytes: ctBase64?.length })
        newSecrets.set(pk.producerId, sharedSecret)
        newSent.add(pk.producerId)
        persistSecrets(myId, epochId, deviceId, newSecrets)
        persistSent(myId, epochId, deviceId, newSent)
      }

      secretsRef.current = newSecrets
      sentToRef.current = newSent
      setSharedSecrets(newSecrets)
      setSentTo(newSent)
      log('encapsulation done', { sentTo: [...newSent], secretPartners: [...newSecrets.keys()] })
    } catch (e) {
      fail('encapsulation failed', e)
      setEncapError(e.message)
    } finally {
      setEncapBusy(false)
    }
  }, [instance, account, myId, epochId, deviceId, partnerKeys])

  // ── Step 3: decapsulate received ciphertexts (re-run if blob changed) ───────
  const performDecapsulation = useCallback(async () => {
    const kp = keyPairRef.current
    if (!kp) {
      warn('decapsulation skipped — no local key pair')
      return
    }
    setEncapBusy(true)
    setEncapError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const { ciphertexts } = await api.getCiphertexts(epochId, deviceId, token)
      log('decapsulation start', { receivedFrom: (ciphertexts ?? []).map((c) => c.senderId) })
      const newSecrets = new Map(secretsRef.current)
      const seen = new Map(ctSeenRef.current)

      for (const ct of ciphertexts ?? []) {
        if (seen.get(ct.senderId) === ct.ciphertextBase64 && newSecrets.has(ct.senderId)) {
          log('decapsulation skip (already done)', { sender: ct.senderId })
          continue
        }
        const ss = await decapsulate(ct.ciphertextBase64, kp.dkBase64)
        log('decapsulated ciphertext', { sender: ct.senderId })
        newSecrets.set(ct.senderId, ss)
        seen.set(ct.senderId, ct.ciphertextBase64)
      }

      persistSecrets(myId, epochId, deviceId, newSecrets)
      persistSeen(myId, epochId, deviceId, seen)
      secretsRef.current = newSecrets
      ctSeenRef.current = seen
      setSharedSecrets(newSecrets)
      setCtSeen(seen)
      log('decapsulation done', { secretPartners: [...newSecrets.keys()] })
    } catch (e) {
      fail('decapsulation failed', e)
      setEncapError(e.message)
    } finally {
      setEncapBusy(false)
    }
  }, [instance, account, myId, epochId, deviceId])

  // ── Step 4: load existing submissions ────────────────────────────────────────
  useEffect(() => {
    if (step !== 4 || !myId) return
    ;(async () => {
      try {
        const token = await api.acquireApiToken(instance, account)
        const my = await api.getMySubmissions(epochId, token)
        const done = new Set((my.submissions ?? []).map((s) => `${s.country}|${s.month}|${s.indicator}|${s.segment}`))
        log('loaded submissions', { cells: done.size })
        setSubmittedCells(done)
      } catch (e) {
        fail('load submissions failed', e)
        setLoadError(e.message)
      }
    })()
  }, [step, myId, epochId, instance, account?.homeAccountId])

  // ── Submit CSV ────────────────────────────────────────────────────────────────
  const onCsvPicked = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''
    setSubmitError(null)
    setCsvResult(null)
    setCsvFileName('')
    if (!file) return
    if (!isCsvFile(file)) {
      setCsvResult(emptyCsvResult('File must be a .csv'))
      setCsvFileName(file.name)
      return
    }
    setCsvFileName(file.name)
    try {
      const text = await file.text()
      const months = epochMonths(epoch)
      const result = parseAndValidateCsv(text, months)
      log('csv parsed', {
        ok: result.ok,
        dataRows: result.dataRowCount,
        rows: result.rows.length,
        errors: result.errors.length,
        missingCells: result.missingCells.length,
      })
      setCsvResult(result)
    } catch (e) {
      fail('csv parse failed', e)
      setCsvResult(emptyCsvResult(e.message || 'Failed to read CSV'))
    }
  }

  const downloadSample = async (event) => {
    const months = epochMonths(epoch)
    if (!months) return
    event.preventDefault()
    try {
      const res = await fetch('/sample.csv')
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      const text = rewriteSampleMonths(await res.text(), months)
      const url = URL.createObjectURL(new Blob([text], { type: 'text/csv' }))
      const a = document.createElement('a')
      a.href = url
      a.download = `sample-${months[0]}.csv`
      a.click()
      URL.revokeObjectURL(url)
    } catch (e) {
      fail('sample download failed', e)
    }
  }

  const submitCsv = async () => {
    if (!csvResult?.ok || submitBusy) return
    setSubmitBusy(true)
    setSubmitError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const payload = []
      for (const row of csvResult.rows) {
        const masked = await calculateMaskedValue(
          row.value, row.country, row.month, row.indicator, row.segment, myId, sharedSecrets
        )
        payload.push({
          country: row.country,
          month: row.month,
          indicator: row.indicator,
          segment: row.segment,
          value: masked.toString(),
          epochId,
          signature: 'web-ui',
        })
      }
      log('submit batch', { rows: payload.length, epochId })
      await api.submitMetricsBatch(payload, token)
      const done = new Set(payload.map((s) => `${s.country}|${s.month}|${s.indicator}|${s.segment}`))
      setSubmittedCells((p) => new Set([...p, ...done]))
      log('submit batch ok', { cells: done.size })
      onRefresh?.()
    } catch (e) {
      fail('submit batch failed', e)
      setSubmitError(e.message)
    } finally {
      setSubmitBusy(false)
    }
  }

  // ── Helpers ──────────────────────────────────────────────────────────────────
  const epochPartnerIds = epoch.producerIds ?? []
  const expectedSmallerIds = epochPartnerIds.filter((id) => id < myId)
  const expectedLargerIds = epochPartnerIds.filter((id) => id > myId)
  const epochKeysReady = (status?.isComplete === true)
    && epochPartnerIds.length >= 2
    && (status?.expectedCount ?? 0) >= 2
    && partnerKeys.length >= epochPartnerIds.filter((id) => id !== myId).length
  const allEncapsDone = expectedSmallerIds.length === 0
    || expectedSmallerIds.every((id) => sentTo.has(id) && sharedSecrets.has(id))
  const allDecapsDone = expectedLargerIds.length === 0
    || expectedLargerIds.every((id) => sharedSecrets.has(id))
  const exchangeComplete = status?.isCiphertextExchangeComplete ?? false
  const keyState = keyStatus(keyPair, status, myId)
  const keysMatch = keyState === 'match'
  const keyProblem = recoveryBlocked ? KEY_RECOVERY_BLOCKED_MESSAGE[keyState] ?? null : null

  // ── Auto-progression ─────────────────────────────────────────────────────────

  // Step 1: bring browser and server keys in sync after restore; at most one attempt per fresh poll
  useEffect(() => {
    if (!hydrated || epochInactive || step !== 1 || keyBusy || recoveryBlocked || !myId || !status) return
    if (status.requestedAt < keyChangedAtRef.current || lastRecoveryStatusRef.current === status) return
    if (keyState === 'none' || keyState === 'local-only') {
      lastRecoveryStatusRef.current = status
      log(keyState === 'none' ? 'auto: generate and register key' : 'auto: re-register existing key')
      generateAndRegister()
    } else if (ROTATE_STATES.has(keyState)) {
      lastRecoveryStatusRef.current = status
      log('auto: rotate key', { keyState })
      rotateKey()
    }
  }, [hydrated, step, keyState, keyBusy, recoveryBlocked, epochInactive, myId, status])

  // Step 1 → 2: advance once all epoch partners registered AND we hold the matching private key
  useEffect(() => {
    if (!hydrated || epochInactive || step !== 1 || !keyPair || !keysMatch) return
    if (!status?.isComplete || !status?.registeredPartners?.includes(myId)) return
    log('step 1 → 2 (all keys registered)')
    setStep(2)
  }, [hydrated, step, !!keyPair, keysMatch, status?.isComplete, status?.registeredCount, epochInactive])

  // Step 2: auto-encapsulate once the full epoch key set is present
  useEffect(() => {
    if (!hydrated || epochInactive || step !== 2 || encapBusy || encapError || !epochKeysReady) return
    if (expectedSmallerIds.length === 0 || allEncapsDone) return
    log('auto: encapsulate', { expectedSmallerIds })
    performEncapsulation()
  }, [hydrated, step, epochKeysReady, encapBusy, encapError, allEncapsDone, epochInactive])

  // Step 2 → 3: wait for every smaller epoch partner, not a partial poll snapshot
  useEffect(() => {
    if (!hydrated || epochInactive || step !== 2 || !epochKeysReady) return
    if (expectedSmallerIds.length === 0 || allEncapsDone) {
      log('step 2 → 3 (encapsulation complete)', { expectedSmallerIds, allEncapsDone })
      setStep(3)
    }
  }, [hydrated, step, epochKeysReady, allEncapsDone, expectedSmallerIds.length, epochInactive])

  // Step 3: auto-decapsulate at most once per fresh poll until every secret is derived;
  // a poll that lands while a decapsulation is in flight is picked up once it finishes
  useEffect(() => {
    if (!hydrated || epochInactive || step !== 3 || encapBusy || encapError || !keyPair || allDecapsDone || !status) return
    if (expectedLargerIds.length === 0 || lastDecapStatusRef.current === status) return
    lastDecapStatusRef.current = status
    log('auto: decapsulate', { expectedLargerIds, actualCiphertexts: status.actualCiphertexts })
    performDecapsulation()
  }, [hydrated, step, status, encapBusy, !!keyPair, encapError, allDecapsDone, epochInactive])

  // Step 3 → 4: advance when ciphertext exchange is complete and all secrets derived
  useEffect(() => {
    if (!hydrated || epochInactive || step !== 3) return
    if (epochPartnerIds.length < 2) return
    if (expectedLargerIds.length === 0 && exchangeComplete) {
      log('step 3 → 4 (nothing to receive, exchange complete)')
      setStep(4)
      return
    }
    if (exchangeComplete && allDecapsDone) {
      log('step 3 → 4 (decapsulation complete)')
      setStep(4)
    }
  }, [hydrated, step, exchangeComplete, allDecapsDone, expectedLargerIds.length, epochPartnerIds.length, epochInactive])

  const producerCount = epochPartnerIds.length || status?.expectedCount || 0

  const setupLabel = () => {
    if (keyRotating) return 'Repairing key…'
    if (keyBusy) return 'Generating key…'
    if (recoveryBlocked) return 'Key needs attention'
    if (!status?.registeredPartners?.includes(myId)) return 'Registering…'
    if (!status?.isComplete) return 'Waiting for all producers…'
    if (step === 3 && allDecapsDone && !exchangeComplete) {
      const remaining = status.missingCiphertextSenders?.length ?? 0
      return remaining > 0
        ? `Waiting for ${remaining} partner${remaining === 1 ? '' : 's'} to finish key exchange…`
        : 'Waiting for partners to finish key exchange…'
    }
    return 'Preparing secure session…'
  }

  const renderStatusBar = () => (
    <div className="card" style={{ marginBottom: '1.5rem' }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(0, 1fr))', gap: '1.25rem' }}>
        <div>
          <div className="form-label" style={{ marginBottom: 4 }}>Epoch</div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', flexWrap: 'wrap' }}>
            <span style={{ fontSize: '1.35rem', fontWeight: 700 }}>{epochId}</span>
            {epoch.isClosed && <span className="status-badge pending">Closed</span>}
          </div>
        </div>
        <div>
          <div className="form-label" style={{ marginBottom: 4 }}>Producers</div>
          <div style={{ fontSize: '1.35rem', fontWeight: 700 }}>
            {status ? `${status.registeredCount}/${producerCount}` : producerCount}
          </div>
        </div>
        <div>
          <div className="form-label" style={{ marginBottom: 4 }}>Public key</div>
          <span className={`status-badge ${KEY_STATUS[keyState].tone}`}>{KEY_STATUS[keyState].label}</span>
          <div className="text-muted" style={{ fontSize: '0.8rem', marginTop: 6 }}>
            Browser <code>{fingerprints.local ?? '—'}</code> · Server <code>{fingerprints.server ?? '—'}</code>
          </div>
        </div>
        <div>
          <div className="form-label" style={{ marginBottom: 4 }}>Your OID</div>
          <code className="text-accent" style={{ fontSize: '0.85rem', wordBreak: 'break-all' }}>{myId || '—'}</code>
        </div>
      </div>
    </div>
  )

  const renderSetup = () => {
    const error = keyError || keyProblem || encapError || pollError
    return (
      <div className="card animate-fade-in">
        <div className="card-header">
          <span className="card-icon">⏳</span>
          <h2 className="card-title">{setupLabel()}</h2>
        </div>

        {error && (
          <div className="info-box error" style={{ marginBottom: '1rem' }}>
            ⚠️ {error}
            {encapError && (
              <button
                className="btn btn-secondary"
                onClick={step === 2 ? performEncapsulation : performDecapsulation}
                disabled={encapBusy}
                style={{ marginLeft: '1rem', padding: '0.25rem 0.6rem', fontSize: '0.8rem' }}
              >
                Retry
              </button>
            )}
          </div>
        )}
      </div>
    )
  }

  const renderSubmit = () => {
    const months = epochMonths(epoch) ?? []
    const allDone = submittedCells.size >= CELL_COUNT
    const errors = csvResult?.errors ?? []
    const fileErrors = errors.filter((e) => !e.record)
    const rowErrors = errors.filter((e) => e.record)
    const shownRowErrors = rowErrors.slice(0, ERROR_DISPLAY_CAP)
    const missingCells = csvResult?.missingCells ?? []
    return (
      <div className="card animate-fade-in">
        <div className="card-header">
          <span className="card-icon">📊</span>
          <h2 className="card-title">Submit Data</h2>
          {months.length > 0 && (
            <span className="text-muted" style={{ marginLeft: 'auto', fontSize: '0.875rem' }}>
              {months[0]} – {months[months.length - 1]}
            </span>
          )}
        </div>

        <div className="info-box">
          Upload a CSV of unmasked values. Noise from your shared secrets is applied in the browser before
          submission — the aggregator only sees masked values. Values must be positive integers (no zeros or negatives).
          {months.length > 0 && (
            <div style={{ marginTop: 8 }}>
              Required <code>month_date</code> values: <strong>{months.join(', ')}</strong>. The file must have
              exactly {CELL_COUNT} data rows ({COUNTRIES.length} countries × {months.length} months × {SERIES.length} series),
              one per cell, with no duplicates.
            </div>
          )}
        </div>

        {loadError && (
          <div className="info-box error">
            ⚠️ {loadError}
          </div>
        )}

        {allDone ? (
          <div className="info-box ok">
            ✅ All {CELL_COUNT} cells submitted. Waiting for the remaining producers…
          </div>
        ) : (
          <>
            <div style={{ display: 'flex', gap: '0.75rem', alignItems: 'center', flexWrap: 'wrap', marginBottom: '1rem' }}>
              <a
                className="btn btn-secondary"
                href="/sample.csv"
                download="sample.csv"
                onClick={downloadSample}
                style={{ fontSize: '0.85rem' }}
              >
                Download sample CSV
              </a>
              <label className="btn btn-secondary" style={{ fontSize: '0.85rem', marginBottom: 0, cursor: 'pointer' }}>
                Choose CSV
                <input
                  type="file"
                  accept=".csv,text/csv"
                  onChange={onCsvPicked}
                  disabled={submitBusy}
                  style={{ display: 'none' }}
                />
              </label>
              {csvFileName && <span className="text-muted" style={{ fontSize: '0.85rem' }}>{csvFileName}</span>}
            </div>

            {errors.length > 0 && (
              <div className="info-box error" style={{ maxHeight: 420, overflowY: 'auto' }}>
                <div style={{ fontWeight: 600, marginBottom: 6 }}>CSV validation failed</div>
                {fileErrors.length > 0 && (
                  <ul style={{ margin: '0 0 0.75rem', paddingLeft: '1.2rem' }}>
                    {fileErrors.map((err, i) => <li key={i}>{err.text}</li>)}
                  </ul>
                )}
                {shownRowErrors.length > 0 && (
                  <div style={{ overflowX: 'auto', background: '#fff', borderRadius: 6 }}>
                    <table className="results-table compact">
                      <thead>
                        <tr>
                          <th>Row</th>
                          {HEADERS.map((h) => <th key={h}>{h}</th>)}
                          <th>Problem</th>
                        </tr>
                      </thead>
                      <tbody>
                        {shownRowErrors.map((err, i) => (
                          <Fragment key={i}>
                            <tr>
                              <td>{err.row}</td>
                              {HEADERS.map((h) => (
                                <td key={h} className={err.column === h ? 'cell-invalid' : undefined}>
                                  {err.record ? (err.record[h] === '' ? '(empty)' : err.record[h]) : ''}
                                </td>
                              ))}
                              <td>{err.message}</td>
                            </tr>
                            {err.duplicateOf && (
                              <tr className="muted">
                                <td>{err.duplicateOf.row}</td>
                                {HEADERS.map((h) => <td key={h}>{err.duplicateOf.record[h]}</td>)}
                                <td>First occurrence</td>
                              </tr>
                            )}
                          </Fragment>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
                {rowErrors.length > ERROR_DISPLAY_CAP && (
                  <div style={{ marginTop: 6 }}>and {rowErrors.length - ERROR_DISPLAY_CAP} more row error(s)</div>
                )}
                {missingCells.length > 0 && (
                  <>
                    <div style={{ fontWeight: 600, margin: '0.75rem 0 6px' }}>
                      Missing cells ({missingCells.length})
                    </div>
                    <div style={{ overflowX: 'auto', background: '#fff', borderRadius: 6 }}>
                      <table className="results-table compact">
                        <thead>
                          <tr>
                            <th>month_date</th>
                            <th>country_iso3</th>
                            <th>indicator</th>
                            <th>segment</th>
                          </tr>
                        </thead>
                        <tbody>
                          {missingCells.slice(0, ERROR_DISPLAY_CAP).map((c, i) => (
                            <tr key={i}>
                              <td>{c.month}</td>
                              <td>{c.country}</td>
                              <td>{c.indicator}</td>
                              <td>{c.segment}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {missingCells.length > ERROR_DISPLAY_CAP && (
                      <div style={{ marginTop: 6 }}>and {missingCells.length - ERROR_DISPLAY_CAP} more</div>
                    )}
                  </>
                )}
              </div>
            )}

            {csvResult?.ok && (
              <>
                <div className="info-box ok">
                  Valid file: {CELL_COUNT} cells covering {months.join(', ')}.
                </div>
                <div style={{ overflowX: 'auto', marginBottom: '1rem' }}>
                  <table className="results-table">
                    <thead>
                      <tr>
                        <th>Country</th>
                        <th>Month</th>
                        <th>Indicator</th>
                        <th>Segment</th>
                        <th>Value</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(csvResult.preview ?? []).map((r, i) => (
                        <tr key={i}>
                          <td>{r.country}</td>
                          <td>{r.month}</td>
                          <td>{r.indicator}</td>
                          <td>{r.segment}</td>
                          <td>{formatInt(r.value)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="text-muted" style={{ fontSize: '0.8rem', marginBottom: '1rem' }}>
                  Showing first {csvResult.preview.length} of {CELL_COUNT} rows
                </div>
              </>
            )}

            {submitError && (
              <div className="info-box error">
                ⚠️ {submitError}
              </div>
            )}

            <button
              className="btn btn-primary"
              onClick={submitCsv}
              disabled={!csvResult?.ok || submitBusy}
            >
              {submitBusy ? 'Submitting…' : 'Submit CSV'}
            </button>
          </>
        )}
      </div>
    )
  }

  const renderWaiting = () => (
    <div className="card animate-fade-in">
      <div className="card-header">
        <span className="card-icon">✅</span>
        <h2 className="card-title">Waiting for a new epoch</h2>
      </div>
      <div className="info-box">
        Epoch {epochId} is closed — every producer has submitted. This page will continue when an admin creates a new epoch.
      </div>
    </div>
  )

  const renderNotEligible = () => (
    <div className="card animate-fade-in">
      <div className="card-header">
        <span className="card-icon">🔒</span>
        <h2 className="card-title">Not eligible for this epoch</h2>
      </div>
      <div className="info-box">
        Your account is not listed as a participant in epoch {epochId}. You can view its status, but you cannot exchange keys or submit metrics.
      </div>
    </div>
  )

  const renderAdminReadOnly = () => (
    <div className="card animate-fade-in">
      <div className="card-header">
        <span className="card-icon">🔒</span>
        <h2 className="card-title">View only</h2>
      </div>
      <div className="info-box">
        Admins cannot exchange keys or submit metrics. Use the Admin Panel to follow epoch {epochId}.
      </div>
    </div>
  )

  if (isAdmin) return <div className="animate-fade-in">{renderAdminReadOnly()}</div>

  return (
    <div className="animate-fade-in">
      {renderStatusBar()}
      {epoch.isEligible === false
        ? renderNotEligible()
        : epoch.isClosed
          ? renderWaiting()
          : step < 4 ? renderSetup() : renderSubmit()}
    </div>
  )
}
