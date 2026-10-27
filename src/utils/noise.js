// Noise formula compatible with the C# SecureNoiseGenerator (HMAC-SHA256 based).
// Noise spans the full signed range of the mask width (64 or 128 bits) and masking wraps
// mod 2^bits, so a masked value reveals nothing about the actual one; the wrapped sum
// across partners cancels exactly.

async function deriveNoise(sharedSecretBytes, country, month, indicator, segment, bits) {
  const key = await crypto.subtle.importKey(
    'raw', sharedSecretBytes, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
  )
  const hmac = new Uint8Array(
    await crypto.subtle.sign(
      'HMAC',
      key,
      new TextEncoder().encode(`${country}|${month}|${indicator}|${segment}`)
    )
  )
  // Signed little-endian integer from the first bits/8 bytes
  let seed = 0n
  for (let i = 0; i < bits / 8; i++) seed |= BigInt(hmac[i]) << BigInt(i * 8)
  return BigInt.asIntN(bits, seed)
}

// secretsMap: Map<partnerId, Uint8Array(32)>
// Sign convention: +1 if myId < partnerId (lexicographic), -1 otherwise
// `actual` may be a number, string, or bigint. Returns bigint in the signed range of `bits`:
// 128 for metric values in newer epochs, 64 for older epochs and the Quorum Check.
export async function calculateMaskedValue(actual, country, month, indicator, segment, myId, secretsMap, bits = 64) {
  let masked = typeof actual === 'bigint' ? actual : BigInt(actual)
  for (const [partnerId, ss] of secretsMap) {
    const noise = await deriveNoise(ss, country, month, indicator, segment, bits)
    masked += noise * (myId < partnerId ? 1n : -1n)
  }
  return BigInt.asIntN(bits, masked)
}
