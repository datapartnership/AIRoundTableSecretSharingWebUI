import { useState, useEffect, useCallback, useRef } from 'react'
import { Routes, Route, Navigate, useLocation, useNavigate } from 'react-router-dom'
import { useIsAuthenticated, useMsal } from '@azure/msal-react'
import Home from './pages/Home'
import AllEpochs from './pages/AllEpochs'
import EpochPage from './pages/EpochPage'
import QuorumPage from './pages/QuorumPage'
import AdminPanel from './pages/AdminPanel'
import { ShellContext } from './components/Shell'
import useActiveQuorum from './hooks/useActiveQuorum'
import { loginRequest, apiTokenRequest } from './authConfig'
import * as api from './utils/api'

const ADMIN_GROUP = import.meta.env.VITE_ADMIN_GROUP_ID

// Groups claim is in the access token, not the ID token — parse it client-side for UI gating only
function parseGroups(accessToken) {
  try {
    const payload = JSON.parse(atob(accessToken.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')))
    return payload.groups ?? []
  } catch {
    return []
  }
}

// Name and organization for the nav; the organization comes from the optional companyName claim
function userOf(account) {
  const claims = account?.idTokenClaims ?? {}
  const email = account?.username || claims.preferred_username || claims.email || ''
  return {
    name: account?.name || claims.name || email,
    org: claims.companyName || email.split('@')[1] || '',
  }
}

function App() {
  const isAuthenticated = useIsAuthenticated()
  const { instance, accounts } = useMsal()
  const account = accounts[0]
  // 'unknown' until the access token's groups are inspected; admins never register as participants
  const [role, setRole] = useState('unknown')
  const isAdmin = role === 'admin'

  useEffect(() => {
    if (!isAuthenticated || !account) { setRole('unknown'); return }
    if (!ADMIN_GROUP) {
      console.error('VITE_ADMIN_GROUP_ID is not configured; admin UI gating is disabled.')
      setRole('partner')
      return
    }
    instance.acquireTokenSilent({ ...apiTokenRequest, account })
      .then(r => setRole(parseGroups(r.accessToken).includes(ADMIN_GROUP) ? 'admin' : 'partner'))
      .catch(() => setRole('partner'))
  }, [isAuthenticated, account?.homeAccountId])

  const [registration, setRegistration] = useState({ state: 'idle', error: null })

  const register = useCallback(async () => {
    if (!account) return
    setRegistration({ state: 'pending', error: null })
    try {
      const token = await api.acquireApiToken(instance, account)
      const r = await api.selfRegister(token)
      console.log('[auth] self-register ok', r)
      setRegistration({ state: 'ok', error: null })
    } catch (e) {
      console.error('[auth] self-register failed', e)
      setRegistration({ state: 'error', error: e.message })
    }
  }, [instance, account?.homeAccountId])

  useEffect(() => {
    if (!isAuthenticated || !account || role !== 'partner') { setRegistration({ state: 'idle', error: null }); return }
    register()
  }, [isAuthenticated, account?.homeAccountId, role, register])

  // One Quorum Check is active at a time; its partners see only the Quorum page until everyone has answered
  const quorum = useActiveQuorum({ enabled: isAuthenticated && role !== 'unknown' })
  const navigate = useNavigate()
  const wasLocked = useRef(false)
  useEffect(() => {
    // Every partner answered, or an admin cancelled it: on to data submission
    if (wasLocked.current && !quorum.locked) navigate('/', { replace: true })
    wasLocked.current = quorum.locked
  }, [quorum.locked])

  // On a page change: back to the top, and focus the page heading for screen readers
  const location = useLocation()
  const lastPath = useRef(null)
  useEffect(() => {
    if (lastPath.current !== null && lastPath.current !== location.pathname) {
      window.scrollTo(0, 0)
      document.querySelector('#main h1')?.focus({ preventScroll: true })
    }
    lastPath.current = location.pathname
  }, [location.pathname])

  const shell = {
    signedIn: isAuthenticated,
    user: userOf(account),
    isAdmin,
    roleKnown: role !== 'unknown',
    registration,
    retryRegistration: register,
    quorum,
    signIn: () => instance.loginRedirect(loginRequest),
    signOut: () => instance.logoutRedirect(),
  }

  // Partners wait for the first Quorum Check status so another page doesn't flash before the redirect
  const quorumPending = role === 'unknown' || (role === 'partner' && !quorum.known)
  const gated = (page) => (!isAuthenticated ? page : quorumPending ? null : quorum.locked ? <Navigate to="/quorum" replace /> : page)
  const signedIn = (page) => (isAuthenticated ? gated(page) : <Navigate to="/" replace />)

  return (
    <ShellContext.Provider value={shell}>
      <div id="app">
        <Routes>
          <Route path="/" element={gated(<Home />)} />
          <Route path="/epochs" element={signedIn(<AllEpochs />)} />
          <Route path="/quorum" element={isAuthenticated ? <QuorumPage /> : <Navigate to="/" replace />} />
          <Route path="/api" element={signedIn(<EpochPage usageType="api" />)} />
          <Route path="/chat" element={signedIn(<EpochPage usageType="chat" />)} />
          <Route path="/admin" element={signedIn(role === 'unknown' ? null : isAdmin ? <AdminPanel /> : <Navigate to="/" replace />)} />
          <Route path="*" element={<Navigate to="/" replace />} />
        </Routes>
      </div>
      <div id="live" className="sr-only" role="status" aria-live="polite" />
    </ShellContext.Provider>
  )
}

export default App
