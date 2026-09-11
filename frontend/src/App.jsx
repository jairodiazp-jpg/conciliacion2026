import { useEffect, useState } from "react";
import { APP_VERSION } from "./generated/buildInfo";
import ProcesoUnificadoSection from "./components/ProcesoUnificadoSection";

const API_BASE = (import.meta.env.VITE_API_URL || "http://127.0.0.1:3000").replace(/\/$/, "");

function IconInicio() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <rect x="3" y="3" width="7" height="7" rx="1.6" fill="currentColor" />
      <rect x="14" y="3" width="7" height="7" rx="1.6" fill="currentColor" />
      <rect x="3" y="14" width="7" height="7" rx="1.6" fill="currentColor" />
      <rect x="14" y="14" width="7" height="7" rx="1.6" fill="currentColor" />
    </svg>
  );
}

function IconProcesos() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M4 19V9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
      <path d="M10 19V5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
      <path d="M16 19v-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
      <path d="M22 19V11" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
    </svg>
  );
}

function IconConfig() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="1.8" />
      <path
        d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1Z"
        stroke="currentColor"
        strokeWidth="1.6"
      />
    </svg>
  );
}

function App() {
  const [view, setView] = useState("inicio");

  useEffect(() => {
    document.title = `Calypso Conciliador v${APP_VERSION}`;
  }, []);

  return (
    <div className="app-frame">
      <div className="app-shell">
        <aside className="sidebar">
          <div className="brand">
            <div className="brand-mark" aria-hidden="true">
              <svg width="26" height="26" viewBox="0 0 26 26" fill="none">
                <path d="M4 13 C8 8, 12 8, 16 13 S22 18, 24 13" stroke="#fff" strokeWidth="2" strokeLinecap="round" />
                <path d="M4 17 C8 12, 12 12, 16 17 S22 22, 24 17" stroke="#7EB6FF" strokeWidth="1.8" strokeLinecap="round" />
              </svg>
            </div>
            <div className="brand-copy">
              <strong>CALYPSO</strong>
              <span>CONCILIADOR</span>
            </div>
          </div>

          <nav className="nav-list" aria-label="Navegación principal">
            <button className={`nav-item ${view === "inicio" ? "active" : ""}`} onClick={() => setView("inicio")}>
              <IconInicio />
              Inicio
            </button>
            <button className={`nav-item ${view === "procesos" ? "active" : ""}`} onClick={() => setView("procesos")}>
              <IconProcesos />
              Procesos
            </button>
            <button className={`nav-item ${view === "configuracion" ? "active" : ""}`} onClick={() => setView("configuracion")}>
              <IconConfig />
              Configuración
            </button>
          </nav>

          <div className="sidebar-status">
            <span className="status-dot" aria-hidden="true">✓</span>
            <div>
              <strong>Sistema actualizado</strong>
              <span>v{APP_VERSION}</span>
            </div>
          </div>
        </aside>

        <main className="workspace">
          <h1 className="page-title">Automatización de pagos</h1>
          <ProcesoUnificadoSection apiBase={API_BASE} view={view} />
        </main>
      </div>
    </div>
  );
}

export default App;
