import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App.tsx'
import { ErrorBoundary } from './ErrorBoundary.tsx'
import { registerServiceWorker } from './pwa/registerServiceWorker.ts'
import { installErrorReporting } from './telemetry/reportError.ts'
import './styles/global.css'
import './styles/components.css'

installErrorReporting()
registerServiceWorker()

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
)
