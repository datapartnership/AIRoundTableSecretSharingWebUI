// The partner's own Quorum Check answers and secret composition tags never leave the browser unmasked.
// Stored per period and device; the epoch page reads them to know which metrics need values above 0.
const key = (myId, periodId, deviceId) => `quorum_${myId}_p${periodId}_${deviceId}`

export function loadQuorumLocal(myId, periodId, deviceId) {
  try { return JSON.parse(localStorage.getItem(key(myId, periodId, deviceId)) ?? 'null') } catch { return null }
}

export function saveQuorumLocal(myId, periodId, deviceId, value) {
  localStorage.setItem(key(myId, periodId, deviceId), JSON.stringify(value))
}

/** Map("country|indicator" -> ticked), or null when this browser has no answers. */
export function answersMap(local) {
  if (!local?.answers) return null
  return new Map(Object.entries(local.answers).map(([k, v]) => [k, v === 1]))
}
