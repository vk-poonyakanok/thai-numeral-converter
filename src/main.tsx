import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { initializeIcons } from '@fluentui/react';
import './index.css'
import App from './App.tsx'

/* global Office */

initializeIcons();

Office.onReady(() => {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <App />
    </StrictMode>,
  )
});
