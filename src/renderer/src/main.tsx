import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import './styles.css'

// StrictMode disabled: Univer mounts once and does not tolerate double-invoke in dev.
ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(<App />)
