import { useState, useEffect, useCallback, useMemo, useRef } from 'react'
import { useMsal } from '@azure/msal-react'
import * as api from '../utils/api'
import { generateMlKemKeyPair, encapsulate, decapsulate } from '../utils/crypto'
import {
  getDeviceId, loadLocalCrypto, wipeLocalCrypto, persistKeyPair, persistSecrets, persistSeen, keyStatus, scopeKeyOf,
} from '../utils/localCrypto'
import { SUPPORT_EMAIL } from '../utils/usage'

const POLL_MS = 2000
const TAG = '[keyexchange]'

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
  mismatch: `This browser’s key does not match the key registered on the server, and it can’t be replaced automatically because the key exchange has already started. Contact ${SUPPORT_EMAIL} to ask an admin to recreate it.`,
  'server-only': `This browser has no private key for the key registered on the server, and it can’t be replaced automatically because the key exchange has already started. Contact ${SUPPORT_EMAIL} to ask an admin to recreate it.`,
  'other-device': `Your key was registered from another browser or device, and it can’t be replaced automatically because the key exchange has already started. Continue there, or contact ${SUPPORT_EMAIL} to ask an admin to recreate it.`,
}

const ROTATE_STATES = new Set(['mismatch', 'server-only', 'other-device'])

/**
 * Runs the ML-KEM key exchange for one scope ({ epochId } or { periodId }) automatically:
 * step 1 generate and register a key, step 2 encapsulate for smaller-ID partners, step 3 decapsulate
 * ciphertexts from larger-ID partners, step 4 every pairwise secret is derived.
 *
 * Mount with a key per scope so all state resets when the scope changes.
 *
 * `phase` summarises progress for the UI: 'creating' (no key registered yet), 'waiting-keys' (waiting for
 * partner keys), 'exchanging' (waiting for ciphertexts), 'ready'.
 */
export default function useKeyExchange({ scope, producerIds, active }) {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  // Azure AD OID matches the JWT sub claim the API uses
  const myId = account?.localAccountId ?? ''
  const deviceId = getDeviceId()
  const scopeKey = scopeKeyOf(scope)
  // Stable for effect dependencies while the scope is the same
  const scopeObj = useMemo(() => scope, [scopeKey])

  const [step, setStep] = useState(1)
  const [hydrated, setHydrated] = useState(false)

  const [keyPair, setKeyPair] = useState(null)
  const [keyBusy, setKeyBusy] = useState(false)
  const [keyError, setKeyError] = useState(null)
  const [recoveryBlocked, setRecoveryBlocked] = useState(false)

  const [status, setStatus] = useState(null)
  const [partnerKeys, setPartnerKeys] = useState([])
  const [pollError, setPollError] = useState(null)
  const [lastPollAt, setLastPollAt] = useState(null)

  const [sentTo, setSentTo] = useState(new Set())       // partners the server holds my ciphertext for
  const [received, setReceived] = useState(new Map())   // senderId → ciphertext blob the server holds for me
  const [sharedSecrets, setSharedSecrets] = useState(new Map())
  const [ctSeen, setCtSeen] = useState(new Map())       // senderId → ciphertext blob we already decapped
  const [encapBusy, setEncapBusy] = useState(false)
  const [encapError, setEncapError] = useState(null)

  const keyPairRef = useRef(null)
  const secretsRef = useRef(new Map())
  const sentToRef = useRef(new Set())
  const ctSeenRef = useRef(new Map())
  // Poll responses requested before the last key change describe the old key and must not trigger recovery
  const keyChangedAtRef = useRef(0)
  const lastRecoveryStatusRef = useRef(null)
  const lastDecapStatusRef = useRef(null)
  // Sent lists requested before our last ciphertext post may not include it yet
  const lastCtPostAtRef = useRef(0)
  keyPairRef.current = keyPair
  secretsRef.current = sharedSecrets
  sentToRef.current = sentTo
  ctSeenRef.current = ctSeen

  const applyClearedCrypto = () => {
    warn('local crypto wiped')
    setKeyPair(null)
    setSharedSecrets(new Map())
    setSentTo(new Set())
    setReceived(new Map())
    setCtSeen(new Map())
    setStep(1)
    setEncapError(null)
    setKeyError(null)
  }

  // ── Restore persisted state before any auto-run ──────────────────────────────
  useEffect(() => {
    if (!myId) return
    const local = loadLocalCrypto(myId, scopeKey, deviceId)
    log('restored local crypto', {
      scope: scopeKey,
      hasKeyPair: !!local.keyPair,
      secretPartners: [...local.secrets.keys()],
      seenFrom: [...local.ctSeen.keys()],
    })
    // Cached secrets only count once the server confirms the matching ciphertexts (see poll)
    setKeyPair(local.keyPair)
    setSharedSecrets(local.secrets)
    setCtSeen(local.ctSeen)
    setHydrated(true)
  }, [myId, scopeKey, deviceId])

  // ── Poll key exchange while setup is in progress ─────────────────────────────
  useEffect(() => {
    if (!hydrated || !myId || !active || step > 3) return
    let alive = true

    const poll = async () => {
      try {
        const token = await api.acquireApiToken(instance, account)
        const requestedAt = Date.now()
        const [rawStatus, pk, sent, inbox] = await Promise.all([
          api.getKeyExchangeStatus(scopeObj, deviceId, token),
          api.getPartnerKeys(scopeObj, deviceId, token),
          api.getSentCiphertexts(scopeObj, deviceId, token),
          api.getCiphertexts(scopeObj, deviceId, token),
        ])
        if (!alive) return

        const s = { ...rawStatus, requestedAt }
        const serverSent = (sent.ciphertexts ?? []).map((c) => c.recipientId)
        const serverReceived = new Map((inbox.ciphertexts ?? []).map((c) => [c.senderId, c.ciphertextBase64]))
        setPollError(null)
        setLastPollAt(new Date())
        setStatus(s)
        setPartnerKeys(pk.partnerKeys ?? [])
        setSentTo((prev) => requestedAt < lastCtPostAtRef.current
          ? new Set([...prev, ...serverSent])
          : new Set(serverSent))
        setReceived(serverReceived)

        log('poll', {
          step,
          scope: scopeKey,
          registered: `${s.registeredCount}/${s.expectedCount}`,
          missing: s.missingPartners,
          ciphertexts: `${s.actualCiphertexts}/${s.expectedCiphertexts}`,
          exchangeComplete: s.isCiphertextExchangeComplete,
          secretPartners: [...secretsRef.current.keys()],
          hasLocalKey: !!keyPairRef.current,
          hasServerKey: !!s.myPublicKeyBase64,
        })
      } catch (e) {
        fail('poll failed', e)
        if (alive) setPollError(e.message)
      }
    }

    poll()
    const id = setInterval(poll, POLL_MS)
    return () => { alive = false; clearInterval(id) }
  }, [hydrated, step, myId, scopeKey, active, instance, account?.homeAccountId, deviceId])

  // ── Step 1: generate & register key pair ─────────────────────────────────────
  const generateAndRegister = useCallback(async () => {
    setKeyBusy(true)
    setKeyError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const kp = await generateMlKemKeyPair()
      log('generated new key pair', { myId, scope: scopeKey })
      await api.registerPublicKey(scopeObj, deviceId, kp.ekBase64, token)
      log('public key registered')
      keyChangedAtRef.current = Date.now()
      const stored = { ...kp, registeredAt: new Date().toISOString() }
      persistKeyPair(myId, scopeKey, deviceId, stored)
      setKeyPair(stored)
    } catch (e) {
      fail('key generate/register failed', e)
      setKeyError(e.message)
    } finally {
      setKeyBusy(false)
    }
  }, [instance, account, myId, scopeKey, deviceId])

  // ── Step 1 recovery: replace a missing/mismatched key while nothing depends on it ─
  const rotateKey = useCallback(async () => {
    setKeyBusy(true)
    setKeyError(null)
    try {
      const token = await api.acquireApiToken(instance, account)
      const kp = await generateMlKemKeyPair()
      warn('rotating key', { myId, scope: scopeKey })
      await api.rotatePublicKey(scopeObj, deviceId, kp.ekBase64, token)
      keyChangedAtRef.current = Date.now()
      wipeLocalCrypto(myId, scopeKey, deviceId)
      applyClearedCrypto()
      const stored = { ...kp, registeredAt: new Date().toISOString() }
      persistKeyPair(myId, scopeKey, deviceId, stored)
      setKeyPair(stored)
      log('key rotated')
    } catch (e) {
      if (e.status === 409 && e.code === 'exchange-started') {
        warn('key rotation blocked — exchange already started', { myId, scope: scopeKey })
        setRecoveryBlocked(true)
      } else {
        fail('key rotation failed', e)
        setKeyError(e.message)
      }
    } finally {
      setKeyBusy(false)
    }
  }, [instance, account, myId, scopeKey, deviceId])

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
        if (newSecrets.has(pk.producerId) && newSent.has(pk.producerId)) continue

        if (newSent.has(pk.producerId) && !newSecrets.has(pk.producerId)) {
          throw new Error(
            `A ciphertext for ${pk.producerId} is already on the server, but this browser lost the shared secret. Contact ${SUPPORT_EMAIL} to ask an admin to recreate it.`
          )
        }

        const { ctBase64, sharedSecret } = await encapsulate(pk.publicKeyBase64)
        await api.postCiphertext(scopeObj, deviceId, pk.producerId, pk.deviceId, ctBase64, token)
        lastCtPostAtRef.current = Date.now()
        log('encapsulated + posted ciphertext', { recipient: pk.producerId })
        newSecrets.set(pk.producerId, sharedSecret)
        newSent.add(pk.producerId)
        persistSecrets(myId, scopeKey, deviceId, newSecrets)
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
  }, [instance, account, myId, scopeKey, deviceId, partnerKeys])

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
      const { ciphertexts } = await api.getCiphertexts(scopeObj, deviceId, token)
      const newSecrets = new Map(secretsRef.current)
      const seen = new Map(ctSeenRef.current)

      for (const ct of ciphertexts ?? []) {
        if (seen.get(ct.senderId) === ct.ciphertextBase64 && newSecrets.has(ct.senderId)) continue
        const ss = await decapsulate(ct.ciphertextBase64, kp.dkBase64)
        log('decapsulated ciphertext', { sender: ct.senderId })
        newSecrets.set(ct.senderId, ss)
        seen.set(ct.senderId, ct.ciphertextBase64)
      }

      persistSecrets(myId, scopeKey, deviceId, newSecrets)
      persistSeen(myId, scopeKey, deviceId, seen)
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
  }, [instance, account, myId, scopeKey, deviceId])

  // ── Helpers ──────────────────────────────────────────────────────────────────
  const expectedSmallerIds = producerIds.filter((id) => id < myId)
  const expectedLargerIds = producerIds.filter((id) => id > myId)
  const keysReady = (status?.isComplete === true)
    && producerIds.length >= 2
    && (status?.expectedCount ?? 0) >= 2
    && partnerKeys.length >= producerIds.filter((id) => id !== myId).length
  const allEncapsDone = expectedSmallerIds.length === 0
    || expectedSmallerIds.every((id) => sentTo.has(id) && sharedSecrets.has(id))
  const allDecapsDone = expectedLargerIds.length === 0
    || expectedLargerIds.every((id) => sharedSecrets.has(id) && received.has(id) && ctSeen.get(id) === received.get(id))
  const exchangeComplete = status?.isCiphertextExchangeComplete ?? false
  const keyState = keyStatus(keyPair, status, myId)
  const keysMatch = keyState === 'match'
  const keyProblem = recoveryBlocked ? KEY_RECOVERY_BLOCKED_MESSAGE[keyState] ?? null : null

  // ── Auto-progression ─────────────────────────────────────────────────────────

  // Step 1: bring browser and server keys in sync after restore; at most one attempt per fresh poll
  useEffect(() => {
    if (!hydrated || !active || step !== 1 || keyBusy || recoveryBlocked || !myId || !status) return
    if (status.requestedAt < keyChangedAtRef.current || lastRecoveryStatusRef.current === status) return
    if (keyState === 'none' || keyState === 'local-only') {
      lastRecoveryStatusRef.current = status
      if (keyState === 'local-only') {
        // The server has no key for this browser (e.g. after a database reset that reused this
        // id), so the cached key and secrets can't belong to this exchange
        warn('auto: server has no key for this browser — discarding cached crypto', { scope: scopeKey })
        wipeLocalCrypto(myId, scopeKey, deviceId)
        applyClearedCrypto()
        keyPairRef.current = null
      }
      log('auto: generate and register key')
      generateAndRegister()
    } else if (ROTATE_STATES.has(keyState)) {
      lastRecoveryStatusRef.current = status
      log('auto: rotate key', { keyState })
      rotateKey()
    }
  }, [hydrated, step, keyState, keyBusy, recoveryBlocked, active, myId, status])

  // Step 1 → 2: advance once all partners registered AND we hold the matching private key
  useEffect(() => {
    if (!hydrated || !active || step !== 1 || !keyPair || !keysMatch) return
    if (!status?.isComplete || !status?.registeredPartners?.includes(myId)) return
    log('step 1 → 2 (all keys registered)')
    setStep(2)
  }, [hydrated, step, !!keyPair, keysMatch, status?.isComplete, status?.registeredCount, active])

  // Step 2: auto-encapsulate once the full key set is present
  useEffect(() => {
    if (!hydrated || !active || step !== 2 || encapBusy || encapError || !keysReady) return
    if (expectedSmallerIds.length === 0 || allEncapsDone) return
    log('auto: encapsulate', { expectedSmallerIds })
    performEncapsulation()
  }, [hydrated, step, keysReady, encapBusy, encapError, allEncapsDone, active])

  // Step 2 → 3: wait for every smaller partner, not a partial poll snapshot
  useEffect(() => {
    if (!hydrated || !active || step !== 2 || !keysReady) return
    if (expectedSmallerIds.length === 0 || allEncapsDone) {
      log('step 2 → 3 (encapsulation complete)')
      setStep(3)
    }
  }, [hydrated, step, keysReady, allEncapsDone, expectedSmallerIds.length, active])

  // Step 3: auto-decapsulate at most once per fresh poll until every secret is derived;
  // a poll that lands while a decapsulation is in flight is picked up once it finishes
  useEffect(() => {
    if (!hydrated || !active || step !== 3 || encapBusy || encapError || !keyPair || allDecapsDone || !status) return
    if (expectedLargerIds.length === 0 || lastDecapStatusRef.current === status) return
    lastDecapStatusRef.current = status
    log('auto: decapsulate', { expectedLargerIds })
    performDecapsulation()
  }, [hydrated, step, status, encapBusy, !!keyPair, encapError, allDecapsDone, active])

  // Step 3 → 4: advance when ciphertext exchange is complete and all secrets derived
  useEffect(() => {
    if (!hydrated || !active || step !== 3) return
    if (producerIds.length < 2) return
    if ((expectedLargerIds.length === 0 || allDecapsDone) && exchangeComplete) {
      log('step 3 → 4 (exchange complete)')
      setStep(4)
    }
  }, [hydrated, step, exchangeComplete, allDecapsDone, expectedLargerIds.length, producerIds.length, active])

  const retry = useCallback(() => {
    setKeyError(null)
    if (step === 2) performEncapsulation()
    else if (step === 3) performDecapsulation()
    else setEncapError(null)
  }, [step, performEncapsulation, performDecapsulation])

  const registered = !!status?.registeredPartners?.includes(myId) && keysMatch
  const phase = step >= 4 ? 'ready'
    : !registered || keyBusy ? 'creating'
    : !status?.isComplete ? 'waiting-keys'
    : 'exchanging'

  return {
    myId,
    hydrated,
    step,
    phase,
    ready: step >= 4,
    keyPair,
    sharedSecrets,
    status,
    error: keyError || keyProblem || encapError,
    pollError,
    lastPollAt,
    retry,
  }
}
