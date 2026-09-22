import { useState, useEffect, useCallback } from 'react'
import { Routes, Route, NavLink, Navigate } from 'react-router-dom'
import { useIsAuthenticated, useMsal } from '@azure/msal-react'
import Home from './pages/Home'
import AdminPanel from './pages/AdminPanel'
import { loginRequest, apiTokenRequest } from './authConfig'
import * as api from './utils/api'
import { APP_VERSION } from './version'

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

  const handleLogin = () => instance.loginRedirect(loginRequest)
  const handleLogout = () => instance.logoutRedirect()
  const email = account?.username
    || account?.idTokenClaims?.preferred_username
    || account?.idTokenClaims?.email
    || ''

  return (
    <div className="app">
      <nav className="navbar">
        <div className="navbar-content">
          <NavLink to="/" className="logo">
            <span className="logo-icon">🔐</span>
            AI Roundtable
            <span className="app-version">{APP_VERSION}</span>
          </NavLink>
          <div className="nav-links">
            <NavLink to="/" end className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}>
              Home
            </NavLink>
            {isAdmin && (
              <NavLink to="/admin" className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}>
                Admin
              </NavLink>
            )}
          </div>
          <div className="nav-auth">
            {isAuthenticated ? (
              <>
                {registration.state === 'pending' && <span className="status-badge pending">Registering…</span>}
                {registration.state === 'ok' && <span className="status-badge success">Registered</span>}
                {registration.state === 'error' && (
                  <button
                    className="status-badge error"
                    onClick={register}
                    title={registration.error ?? ''}
                    style={{ border: 'none', cursor: 'pointer' }}
                  >
                    Registration failed · Retry
                  </button>
                )}
                <span className="nav-user" title={email}>{email}</span>
                <button className="btn btn-secondary nav-btn" onClick={handleLogout}>Sign out</button>
              </>
            ) : (
              <button className="btn btn-primary nav-btn" onClick={handleLogin}>Sign in</button>
            )}
          </div>
        </div>
      </nav>

      <main className="main-content">
        <Routes>
          <Route path="/" element={<Home isAdmin={isAdmin} roleKnown={role !== 'unknown'} />} />
          <Route path="/admin" element={isAuthenticated ? <AdminPanel /> : <Navigate to="/" replace />} />
          <Route path="*" element={<Navigate to="/" replace />} />
        </Routes>
      </main>

      <footer className="site-footer">
        <div className="footer-bottom">
          <span>AI Roundtable</span>
          <span className="app-version">{APP_VERSION}</span>
        </div>
      </footer>
    </div>
  )
}

export default App
