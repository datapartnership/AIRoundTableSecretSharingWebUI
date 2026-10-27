import { useCallback, useEffect, useState } from 'react'
import { useMsal } from '@azure/msal-react'
import * as api from '../utils/api'

const LOCKED_POLL_MS = 5000
const IDLE_POLL_MS = 20000

/** True when the caller is a partner in the active Quorum Check and not everyone has answered. */
export const isQuorumLocked = (status) =>
  !!(status?.period && status.isEligible && !status.quorumComplete && !status.period.isCancelled)

/**
 * The one active Quorum Check, polled so a new or replaced period shows up without a reload.
 * Polls faster while the caller is locked to it, and refreshes when the tab regains focus.
 */
export default function useActiveQuorum({ enabled }) {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  const [status, setStatus] = useState(null)
  const [known, setKnown] = useState(false)
  const [error, setError] = useState(null)

  const refresh = useCallback(async () => {
    try {
      const token = await api.acquireApiToken(instance, account)
      setStatus(await api.getQuorumCurrent(token))
      setError(null)
    } catch (e) {
      console.error('[quorum] active status failed', e)
      setError(e.message)
    } finally {
      setKnown(true)
    }
  }, [instance, account?.homeAccountId])

  const locked = enabled && isQuorumLocked(status)

  useEffect(() => {
    if (!enabled || !account) { setStatus(null); setKnown(false); setError(null); return }
    refresh()
  }, [enabled, account?.homeAccountId, refresh])

  useEffect(() => {
    if (!enabled || !account) return
    const id = setInterval(refresh, locked ? LOCKED_POLL_MS : IDLE_POLL_MS)
    const onVisible = () => { if (document.visibilityState === 'visible') refresh() }
    document.addEventListener('visibilitychange', onVisible)
    window.addEventListener('focus', refresh)
    return () => {
      clearInterval(id)
      document.removeEventListener('visibilitychange', onVisible)
      window.removeEventListener('focus', refresh)
    }
  }, [enabled, account?.homeAccountId, refresh, locked])

  return { status, known, error, refresh, locked }
}
