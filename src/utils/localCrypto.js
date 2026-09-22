import { bytesToBase64, base64ToBytes } from './crypto'

const deviceKey = 'mlkem_device_id'

export function getDeviceId() {
  let id = localStorage.getItem(deviceKey)
  if (!id) {
    id = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`
    localStorage.setItem(deviceKey, id)
  }
  return id
}

const scoped = (id, epochId, deviceId) => `${id}_${epochId}_${deviceId}`
export const kpKey = (id, epochId, deviceId) => `mlkem_kp_${scoped(id, epochId, deviceId)}`
export const ssKey = (id, epochId, deviceId) => `mlkem_ss_${scoped(id, epochId, deviceId)}`
export const ctKey = (id, epochId, deviceId) => `mlkem_ct_${scoped(id, epochId, deviceId)}`
export const ctSeenKey = (id, epochId, deviceId) => `mlkem_ctseen_${scoped(id, epochId, deviceId)}`

export function wipeLocalCrypto(id, epochId, deviceId) {
  localStorage.removeItem(kpKey(id, epochId, deviceId))
  localStorage.removeItem(ssKey(id, epochId, deviceId))
  localStorage.removeItem(ctKey(id, epochId, deviceId))
  localStorage.removeItem(ctSeenKey(id, epochId, deviceId))
}

export function persistKeyPair(id, epochId, deviceId, kp) {
  localStorage.setItem(kpKey(id, epochId, deviceId), JSON.stringify(kp))
}

export function persistSecrets(id, epochId, deviceId, map) {
  localStorage.setItem(ssKey(id, epochId, deviceId), JSON.stringify(
    Object.fromEntries([...map].map(([k, v]) => [k, bytesToBase64(v)]))
  ))
}

export function persistSent(id, epochId, deviceId, sent) {
  localStorage.setItem(ctKey(id, epochId, deviceId), JSON.stringify([...sent]))
}

export function persistSeen(id, epochId, deviceId, map) {
  localStorage.setItem(ctSeenKey(id, epochId, deviceId), JSON.stringify(Object.fromEntries(map)))
}

export function loadLocalKeyPair(id, epochId, deviceId) {
  const kp = localStorage.getItem(kpKey(id, epochId, deviceId))
  if (!kp) return null
  try { return JSON.parse(kp) } catch { return null }
}

export function loadLocalCrypto(id, epochId, deviceId) {
  const keyPair = loadLocalKeyPair(id, epochId, deviceId)
  const secrets = new Map()
  let sentTo = new Set()
  const ctSeen = new Map()

  const ss = localStorage.getItem(ssKey(id, epochId, deviceId))
  if (ss) {
    try {
      for (const [k, v] of Object.entries(JSON.parse(ss))) {
        secrets.set(k, base64ToBytes(v))
      }
    } catch { /* ignore corrupt cache */ }
  }

  const ct = localStorage.getItem(ctKey(id, epochId, deviceId))
  if (ct) {
    try { sentTo = new Set(JSON.parse(ct)) } catch { sentTo = new Set() }
  }

  const seen = localStorage.getItem(ctSeenKey(id, epochId, deviceId))
  if (seen) {
    try {
      for (const [k, v] of Object.entries(JSON.parse(seen))) ctSeen.set(k, v)
    } catch { /* ignore corrupt cache */ }
  }

  return { keyPair, secrets, sentTo, ctSeen }
}

export const KEY_STATUS = {
  match: { label: 'Key matches server', tone: 'success' },
  mismatch: { label: 'Mismatch', tone: 'error' },
  'server-only': { label: 'Server has key, browser has none', tone: 'error' },
  'local-only': { label: 'Not registered yet', tone: 'pending' },
  none: { label: 'Not registered yet', tone: 'pending' },
  'other-device': { label: 'Registered from another device', tone: 'error' },
  unknown: { label: 'Checking key…', tone: 'pending' },
}

// serverStatus is the GET /keyexchange/status response for this device, or null if not loaded yet.
export function keyStatus(localKp, serverStatus, myId) {
  if (!serverStatus) return 'unknown'
  const serverKey = serverStatus.myPublicKeyBase64
  if (serverKey && localKp) return serverKey === localKp.ekBase64 ? 'match' : 'mismatch'
  if (serverKey) return 'server-only'
  if (serverStatus.registeredPartners?.includes(myId)) return 'other-device'
  return localKp ? 'local-only' : 'none'
}

export async function keyFingerprint(ekBase64) {
  if (!ekBase64) return null
  const digest = await globalThis.crypto.subtle.digest('SHA-256', new TextEncoder().encode(ekBase64))
  return [...new Uint8Array(digest)].slice(0, 4).map((b) => b.toString(16).padStart(2, '0')).join('')
}
