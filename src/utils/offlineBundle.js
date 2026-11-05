// Offline submission for partners whose data computer has no internet access.
//
// 1. Online, the epoch page exports a key file: for every cell the partner owes, the net noise
//    (the "pad") their pairwise secrets add to it, encrypted with a passphrase. The secrets
//    themselves never leave the browser.
// 2. Offline, offline-masker.html decrypts the key file, checks the CSV and writes a masked file:
//    masked = (value + pad) wrapped to the epoch's mask width, the same as calculateMaskedValue.
// 3. Online, the epoch page checks the masked file and submits it.
//
// No MSAL or API imports: the offline tool bundles this module.
import { calculateMaskedValue } from './noise'
import { cellKey, maxValue, requiredCells } from './csvUpload'

export const KEY_FILE_FORMAT = 'airt-offline-key'
export const MASKED_FILE_FORMAT = 'airt-masked'
export const OFFLINE_FORMAT_VERSION = 1
export const MIN_PASSPHRASE = 12
export const OFFLINE_TOOL_FILE = 'offline-masker.html'
export const OFFLINE_SIGNATURE = 'web-ui-offline'
const PBKDF2_ITERATIONS = 600_000

// Local copies of the crypto.js helpers: importing crypto.js would bundle ML-KEM into the offline tool
function bytesToBase64(bytes) {
  let bin = ''
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i])
  return btoa(bin)
}

function base64ToBytes(b64) {
  const bin = atob(b64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return bytes
}

const subtle = () => {
  if (!globalThis.crypto?.subtle) throw new Error('This browser has no Web Crypto support. Use a current version of Edge, Chrome or Firefox.')
  return globalThis.crypto.subtle
}
const randomBytes = (n) => globalThis.crypto.getRandomValues(new Uint8Array(n))
const randomId = () => globalThis.crypto.randomUUID?.() ?? bytesToBase64(randomBytes(16))

/** SHA-256 of the public key, hex: ties both files to the key this browser registered for the epoch. */
export async function fullKeyFingerprint(ekBase64) {
  const digest = await subtle().digest('SHA-256', new TextEncoder().encode(ekBase64))
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('')
}

async function deriveAesKey(passphrase, salt, iterations) {
  const base = await subtle().importKey('raw', new TextEncoder().encode(passphrase), 'PBKDF2', false, ['deriveKey'])
  return subtle().deriveKey(
    { name: 'PBKDF2', hash: 'SHA-256', salt, iterations },
    base,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt'],
  )
}

/**
 * Builds the key file content for an epoch (plain object; serialise with JSON.stringify).
 * `quorumRules` = { answers: Map | null, ignoredPairs: Set } as used by parseAndValidateCsv.
 */
export async function buildKeyFile({ epoch, months, maskBits, myId, keyPair, sharedSecrets, quorumRules, passphrase, epochLabel }) {
  if ((passphrase ?? '').length < MIN_PASSPHRASE) throw new Error(`The passphrase needs at least ${MIN_PASSPHRASE} characters.`)
  const pads = {}
  for (const c of requiredCells(months, quorumRules)) {
    const pad = await calculateMaskedValue(0n, c.country, c.month, c.indicator, c.segment, myId, sharedSecrets, maskBits)
    pads[cellKey(c.country, c.month, c.indicator, c.segment)] = pad.toString()
  }
  const payload = {
    epochId: epoch.epochId,
    epochLabel: epochLabel ?? null,
    usageType: epoch.usageType ?? null,
    months,
    maskBits,
    myId,
    keyFingerprint: await fullKeyFingerprint(keyPair.ekBase64),
    keyId: randomId(),
    createdAt: new Date().toISOString(),
    quorum: {
      answers: quorumRules.answers ? Object.fromEntries([...quorumRules.answers].map(([k, v]) => [k, v ? 1 : 0])) : null,
      ignoredPairs: [...(quorumRules.ignoredPairs ?? [])],
    },
    pads,
  }
  const salt = randomBytes(16)
  const iv = randomBytes(12)
  const key = await deriveAesKey(passphrase, salt, PBKDF2_ITERATIONS)
  const ciphertext = new Uint8Array(await subtle().encrypt({ name: 'AES-GCM', iv }, key, new TextEncoder().encode(JSON.stringify(payload))))
  return {
    format: KEY_FILE_FORMAT,
    v: OFFLINE_FORMAT_VERSION,
    // Not secret: lets the offline tool say which epoch the file is for before the passphrase is entered
    epochId: epoch.epochId,
    epochLabel: epochLabel ?? null,
    createdAt: payload.createdAt,
    kdf: { name: 'PBKDF2', hash: 'SHA-256', iterations: PBKDF2_ITERATIONS, salt: bytesToBase64(salt) },
    cipher: { name: 'AES-GCM', iv: bytesToBase64(iv) },
    ciphertext: bytesToBase64(ciphertext),
  }
}

/** Parses the outer key file without the passphrase. Throws with a user-facing message. */
export function readKeyFileHeader(text) {
  let file
  try { file = JSON.parse(text) } catch { throw new Error('This is not a key file. Choose the .json key file you exported from the epoch page.') }
  if (file?.format !== KEY_FILE_FORMAT) throw new Error('This is not a key file. Choose the .json key file you exported from the epoch page.')
  if (file.v !== OFFLINE_FORMAT_VERSION) throw new Error(`This key file has version ${file.v}; this tool reads version ${OFFLINE_FORMAT_VERSION}. Download the offline tool again from the epoch page.`)
  if (file.kdf?.name !== 'PBKDF2' || file.cipher?.name !== 'AES-GCM' || !file.ciphertext) throw new Error('The key file is damaged. Export it again from the epoch page.')
  return file
}

/** Decrypts a key file into { epochId, months, maskBits, myId, keyFingerprint, keyId, quorumRules, pads: Map<cellKey, bigint> }. */
export async function openKeyFile(text, passphrase) {
  const file = readKeyFileHeader(text)
  let plain
  try {
    const key = await deriveAesKey(passphrase, base64ToBytes(file.kdf.salt), file.kdf.iterations)
    plain = await subtle().decrypt({ name: 'AES-GCM', iv: base64ToBytes(file.cipher.iv) }, key, base64ToBytes(file.ciphertext))
  } catch {
    throw new Error('The passphrase is wrong, or the key file is damaged. Check the passphrase and try again.')
  }
  const p = JSON.parse(new TextDecoder().decode(plain))
  const pads = new Map(Object.entries(p.pads).map(([k, v]) => [k, BigInt(v)]))
  return {
    epochId: p.epochId,
    epochLabel: p.epochLabel,
    usageType: p.usageType,
    months: p.months,
    maskBits: p.maskBits,
    myId: p.myId,
    keyFingerprint: p.keyFingerprint,
    keyId: p.keyId,
    createdAt: p.createdAt,
    quorumRules: {
      answers: p.quorum?.answers ? new Map(Object.entries(p.quorum.answers).map(([k, v]) => [k, v === 1])) : null,
      ignoredPairs: new Set(p.quorum?.ignoredPairs ?? []),
    },
    pads,
  }
}

/** Masks validated CSV rows with an opened key file and returns the masked file content (plain object). */
export function buildMaskedFile(keyFile, rows, toolVersion) {
  const values = rows.map((r) => {
    const key = cellKey(r.country, r.month, r.indicator, r.segment)
    const pad = keyFile.pads.get(key)
    if (pad === undefined) throw new Error(`The key file has no mask for ${key}.`)
    const masked = BigInt.asIntN(keyFile.maskBits, BigInt(r.value) + pad)
    return { country: r.country, month: r.month, indicator: r.indicator, segment: r.segment, value: masked.toString() }
  })
  return {
    format: MASKED_FILE_FORMAT,
    v: OFFLINE_FORMAT_VERSION,
    epochId: keyFile.epochId,
    epochLabel: keyFile.epochLabel,
    myId: keyFile.myId,
    maskBits: keyFile.maskBits,
    keyFingerprint: keyFile.keyFingerprint,
    keyId: keyFile.keyId,
    createdAt: new Date().toISOString(),
    toolVersion: toolVersion ?? null,
    values,
  }
}

/**
 * Checks a masked file against the epoch page's state.
 * Returns { ok: true, rows, file } or { ok: false, problems: string[] }.
 */
export async function checkMaskedFile(text, { epochId, myId, maskBits, keyPair, months, quorumRules }) {
  let file
  try { file = JSON.parse(text) } catch { return { ok: false, problems: ['This is not a masked file. Choose the .json file the offline tool created.'] } }
  if (file?.format !== MASKED_FILE_FORMAT) return { ok: false, problems: ['This is not a masked file. Choose the .json file the offline tool created.'] }
  if (file.v !== OFFLINE_FORMAT_VERSION) return { ok: false, problems: [`This masked file has version ${file.v}; this page reads version ${OFFLINE_FORMAT_VERSION}. Download the offline tool again and mask your data again.`] }

  const problems = []
  if (file.epochId !== epochId) problems.push(`The file is for another epoch (epoch ${file.epochId}${file.epochLabel ? `, ${file.epochLabel}` : ''}). Open the epoch it was made for, or export a key file for this epoch and mask your data again.`)
  if (file.myId !== myId) problems.push('The file was made with a key file of another user. Sign in as the user who exported the key file.')
  if (file.maskBits !== maskBits) problems.push(`The file uses ${file.maskBits}-bit masks; this epoch uses ${maskBits}-bit masks. Export a new key file and mask your data again.`)
  if (!keyPair) problems.push('This browser has no key for this epoch. Submit from the browser that exported the key file.')
  else if (file.keyFingerprint !== await fullKeyFingerprint(keyPair.ekBase64)) {
    problems.push('The file was masked with a key that does not match this browser’s key for this epoch. Submit from the browser that exported the key file, or export a new key file here and mask your data again.')
  }
  if (problems.length) return { ok: false, problems }

  const values = Array.isArray(file.values) ? file.values : []
  const required = new Set(requiredCells(months, quorumRules).map((c) => cellKey(c.country, c.month, c.indicator, c.segment)))
  const min = -maxValue(maskBits) - 1n
  const max = maxValue(maskBits)
  const seen = new Set()
  const rows = []
  let extra = 0, duplicate = 0, bad = 0
  for (const v of values) {
    const key = cellKey(v?.country, v?.month, v?.indicator, v?.segment)
    if (!required.has(key)) { extra++; continue }
    if (seen.has(key)) { duplicate++; continue }
    seen.add(key)
    let masked
    try { masked = BigInt(v.value) } catch { bad++; continue }
    if (typeof v.value !== 'string' || masked < min || masked > max) { bad++; continue }
    rows.push({ country: v.country, month: v.month, indicator: v.indicator, segment: v.segment, masked })
  }
  const missing = [...required].filter((k) => !seen.has(k)).length
  if (extra) problems.push(`${extra} values are for cells this epoch does not expect${quorumRules.ignoredPairs?.size ? ', or for metrics below the quorum' : ''}.`)
  if (duplicate) problems.push(`${duplicate} values repeat a cell.`)
  if (bad) problems.push(`${bad} values are not valid ${maskBits}-bit masked numbers.`)
  if (missing) problems.push(`${missing} of the ${required.size} expected values are missing.`)
  if (problems.length) {
    problems.push('Do not edit the masked file. Check your CSV in the offline tool and mask it again.')
    return { ok: false, problems }
  }
  return { ok: true, rows, file }
}
