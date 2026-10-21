import { InteractionRequiredAuthError } from '@azure/msal-browser'
import { apiTokenRequest } from '../authConfig'

const BASE = `${(import.meta.env.VITE_API_URL ?? '').replace(/\/$/, '')}/api`

export async function acquireApiToken(msalInstance, account) {
  try {
    const r = await msalInstance.acquireTokenSilent({ ...apiTokenRequest, account })
    return r.accessToken
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const r = await msalInstance.acquireTokenPopup({ ...apiTokenRequest, account })
      return r.accessToken
    }
    throw e
  }
}

function hdr(token, json = false) {
  const h = { Authorization: `Bearer ${token}` }
  if (json) h['Content-Type'] = 'application/json'
  return h
}

async function get(path, token) {
  const r = await fetch(`${BASE}${path}`, { headers: hdr(token) })
  if (!r.ok) {
    const body = await r.json().catch(() => ({}))
    throw Object.assign(new Error(body.error || body.message || r.statusText), { status: r.status, code: body.code })
  }
  return r.json()
}

async function post(path, body, token) {
  const r = await fetch(`${BASE}${path}`, {
    method: 'POST', headers: hdr(token, true), body: JSON.stringify(body),
  })
  if (!r.ok) {
    const data = await r.json().catch(() => ({}))
    throw Object.assign(new Error(data.error || data.message || r.statusText), { status: r.status, code: data.code })
  }
  return r.json()
}

// ── Registry ──────────────────────────────────────────────────────────────────
export const getProducers = (token) => get('/registry/producers', token)
export const getEpochs = (token) => get('/registry/epochs', token)
export const selfRegister = (token) => post('/registry/producers/me', {}, token)

// ── Key Exchange ──────────────────────────────────────────────────────────────
// `scope` is { epochId } for an epoch or { periodId } for a Quorum Check period.
const scopeQuery = (scope, deviceId) =>
  new URLSearchParams({ ...Object.fromEntries(Object.entries(scope).map(([k, v]) => [k, String(v)])), deviceId }).toString()

export const registerPublicKey = (scope, deviceId, publicKeyBase64, token) =>
  post('/keyexchange/register', { ...scope, deviceId, publicKeyBase64 }, token)

export const rotatePublicKey = (scope, deviceId, publicKeyBase64, token) =>
  post('/keyexchange/rotate', { ...scope, deviceId, publicKeyBase64 }, token)

export const getPartnerKeys = (scope, deviceId, token) =>
  get(`/keyexchange/keys?${scopeQuery(scope, deviceId)}`, token)

export const getKeyExchangeStatus = (scope, deviceId, token) =>
  get(`/keyexchange/status?${scopeQuery(scope, deviceId)}`, token)

// ── Ciphertexts ───────────────────────────────────────────────────────────────
export const postCiphertext = (scope, deviceId, recipientId, recipientDeviceId, ciphertextBase64, token) =>
  post('/ciphertext', { ...scope, deviceId, recipientId, recipientDeviceId, ciphertextBase64 }, token)

export const getCiphertexts = (scope, deviceId, token) =>
  get(`/ciphertext?${scopeQuery(scope, deviceId)}`, token)

export const getSentCiphertexts = (scope, deviceId, token) =>
  get(`/ciphertext/sent?${scopeQuery(scope, deviceId)}`, token)

// ── Metrics ───────────────────────────────────────────────────────────────────
export const getMySubmissions = (epochId, token) =>
  get(`/metrics/mysubmissions?epochId=${encodeURIComponent(epochId)}`, token)

export const submitMetricsBatch = (submissions, token) =>
  post('/metrics/submit-batch', submissions, token)

// ── Quorum Check ──────────────────────────────────────────────────────────────
export const getQuorumCurrent = (usageType, token) =>
  get(`/quorum/current?usageType=${encodeURIComponent(usageType)}`, token)

export const getQuorumPeriod = (periodId, token) =>
  get(`/quorum/periods/${encodeURIComponent(periodId)}`, token)

export const submitQuorum = (periodId, answers, token) =>
  post('/quorum/submit', { periodId, answers }, token)

// ── Admin ─────────────────────────────────────────────────────────────────────
export const adminReset = (token) => post('/admin/reset', {}, token)

export const adminResetAndCreateEpoch = (body, token) =>
  post('/admin/producers/reset-and-create-epoch', body, token)

export const adminGetEpochs = (token) => get('/admin/epochs', token)

export const adminGetEpochDetail = (epochId, token) => get(`/admin/epochs/${epochId}`, token)

export const adminCancelEpoch = (epochId, reason, token) =>
  post(`/admin/epochs/${epochId}/cancel`, { reason }, token)

export const adminRecreateEpoch = (epochId, reason, token) =>
  post(`/admin/epochs/${epochId}/recreate`, { reason }, token)

export const adminGetQuorumPeriods = (token) => get('/admin/quorum-periods', token)

export const adminCreateQuorumPeriod = (body, token) => post('/admin/quorum-periods', body, token)

export const adminCancelQuorumPeriod = (periodId, reason, token) =>
  post(`/admin/quorum-periods/${periodId}/cancel`, { reason }, token)
