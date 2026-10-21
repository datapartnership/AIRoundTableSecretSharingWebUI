import { useCallback, useEffect, useState } from 'react'
import { useMsal } from '@azure/msal-react'
import * as api from '../utils/api'

const POLL_MS = 5000

/** The caller's epochs (open, plus closed ones they took part in), refreshed while the page is open. */
export default function useEpochs({ poll = true } = {}) {
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  const [epochs, setEpochs] = useState(null)
  const [error, setError] = useState(null)

  const refresh = useCallback(async () => {
    try {
      const token = await api.acquireApiToken(instance, account)
      const data = await api.getEpochs(token)
      setEpochs(data.epochs ?? [])
      setError(null)
    } catch (e) {
      console.error('[epochs] refresh failed', e)
      setError(e.message)
    }
  }, [instance, account?.homeAccountId])

  useEffect(() => {
    if (!account) return
    refresh()
    if (!poll) return
    const id = setInterval(refresh, POLL_MS)
    return () => clearInterval(id)
  }, [account?.homeAccountId, refresh, poll])

  const retry = useCallback(() => { setError(null); setEpochs(null); refresh() }, [refresh])

  return { epochs, error, refresh, retry, loading: epochs === null && !error }
}

/** Newest first, as in the wireframe lists. */
export const byStartDesc = (a, b) => String(b.startDate).localeCompare(String(a.startDate)) || b.epochId - a.epochId

/** The partner's contribution status for a row: 'todo' needs their data, 'done' submitted. */
export const contribution = (e) => (e.isClosed || e.mySubmissionComplete ? 'done' : 'todo')
