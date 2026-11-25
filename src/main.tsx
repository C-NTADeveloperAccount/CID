import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication } from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import { msalConfig, registeredAppOrigin } from './authConfig'
import App from './App'
import './styles.css'

const msalInstance = new PublicClientApplication(msalConfig)

async function start() {
  const isPopupWindow = Boolean(window.opener)
  const isAuthCallback = window.location.pathname === '/redirect'

  if (!isPopupWindow && window.location.origin !== registeredAppOrigin) {
    window.location.replace(`${registeredAppOrigin}/`)
    return
  }

  await msalInstance.initialize()
  const authResult = await msalInstance.handleRedirectPromise()

  if (authResult?.account) {
    msalInstance.setActiveAccount(authResult.account)
  }

  if (isPopupWindow && isAuthCallback) {
    if (authResult) window.close()
    createRoot(document.getElementById('root')!).render(
      <main className="status-screen">
        <div className="status-icon">✓</div>
        <h1>Sign-in complete</h1>
        <p>You can close this window and return to CID.</p>
      </main>,
    )
    return
  }

  const existingAccount = msalInstance.getAllAccounts()[0]
  if (existingAccount && !msalInstance.getActiveAccount()) {
    msalInstance.setActiveAccount(existingAccount)
  }

  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>,
  )
}

void start()
