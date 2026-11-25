import type { Configuration, PopupRequest } from '@azure/msal-browser'

const clientId = (import.meta.env.VITE_ENTRA_CLIENT_ID || import.meta.env.VITE_MSAL_CLIENT_ID)?.trim()
const tenantId = (import.meta.env.VITE_ENTRA_TENANT_ID || import.meta.env.VITE_MSAL_TENANT_ID)?.trim()
const authority = import.meta.env.VITE_ENTRA_AUTHORITY?.trim()
const redirectUri = (import.meta.env.VITE_REDIRECT_URI || import.meta.env.VITE_MSAL_REDIRECT_URI || window.location.origin).trim()
export const registeredAppOrigin = new URL(redirectUri, window.location.origin).origin

export const isMsalConfigured = Boolean(
  clientId &&
    tenantId &&
    clientId !== '00000000-0000-0000-0000-000000000000' &&
    tenantId !== '00000000-0000-0000-0000-000000000000',
)

export const msalConfig: Configuration = {
  auth: {
    clientId: clientId || '00000000-0000-0000-0000-000000000000',
    authority: authority || `https://login.microsoftonline.com/${tenantId || 'organizations'}`,
    redirectUri,
    postLogoutRedirectUri: registeredAppOrigin,
    navigateToLoginRequestUrl: true,
    clientCapabilities: ['CP1'],
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: true,
  },
}

export const loginRequest: PopupRequest = {
  scopes: ['openid', 'profile', 'email'],
  prompt: 'select_account',
}
