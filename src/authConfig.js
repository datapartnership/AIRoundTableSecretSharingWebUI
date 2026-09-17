import { PublicClientApplication } from '@azure/msal-browser'

const BASE = `${(import.meta.env.VITE_REDIRECT_URL ?? '').replace(/\/$/, '')}/`
const apiScope = import.meta.env.VITE_MSAL_API_SCOPE

if (!apiScope) {
  throw new Error('VITE_MSAL_API_SCOPE must be configured with the backend delegated API scope.')
}

// Optional, comma-separated list of additional authorities MSAL should trust
// (e.g. "localhost:5100"). Not set by default; only used when configured.
const knownAuthorities = (import.meta.env.VITE_MSAL_KNOWN_AUTHORITIES ?? '')
  .split(',')
  .map((authority) => authority.trim())
  .filter(Boolean)

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_MSAL_CLIENT_ID,
    authority: import.meta.env.VITE_MSAL_AUTHORITY,
    redirectUri: BASE,
    ...(knownAuthorities.length > 0 ? { knownAuthorities } : {}),
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
}

export const loginRequest = {
  scopes: ['openid', 'profile', 'email', apiScope],
}

export const apiTokenRequest = {
  scopes: [apiScope],
}

export const msalInstance = new PublicClientApplication(msalConfig)
