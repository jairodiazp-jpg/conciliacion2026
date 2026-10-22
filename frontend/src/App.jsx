import { useEffect, useState } from "react";
import { APP_VERSION } from "./generated/buildInfo";
import ProcesoUnificadoSection from "./components/ProcesoUnificadoSection";
import QueryMemorandoSection from "./components/QueryMemorandoSection";
import VoucherSearchSection from "./components/VoucherSearchSection";

const configuredApi = String(import.meta.env.VITE_API_URL || "").replace(/\/$/, "");
const API_BASE = import.meta.env.DEV ? "" : configuredApi;

function IconInicio() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="m4 10 8-6 8 6v9a1 1 0 0 1-1 1h-5v-6h-4v6H5a1 1 0 0 1-1-1v-9Z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    </svg>
  );
}

function IconProcesos() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M4 19V9M10 19V5M16 19v-6M22 19V11" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
      <path d="m3 5 4-2 4 3 5-3 5 2" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
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

function IconQuery() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M6 3h8l4 4v14H6a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" />
      <path d="M14 3v5h5M8 12h6M8 16h4" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
      <circle cx="16.5" cy="16.5" r="2.5" stroke="currentColor" strokeWidth="1.7" />
      <path d="m18.3 18.3 2 2" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconVoucher() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <rect x="4" y="6" width="16" height="13" rx="3" stroke="currentColor" strokeWidth="1.7" />
      <path d="M8 6V4h8v2M8 12h.01M16 12h.01M9 16h6" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
      <path d="M9 2v2M15 2v2" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
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
            <img src="/calypso-logo.svg" alt="CALYPSO Conciliador" className="brand-logo" />
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
            <button className={`nav-item ${view === "query-memorando" ? "active" : ""}`} onClick={() => setView("query-memorando")}>
              <IconQuery />
              Conciliar Query y memorando ya conciliado
            </button>
            <button className={`nav-item ${view === "voucher" ? "active" : ""}`} onClick={() => setView("voucher")}>
              <IconVoucher />
              Agente para buscar valores en voucher
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
          <h1 className="page-title">
            {view === "query-memorando" ? "Conciliar Query y memorando ya conciliado" : view === "voucher" ? "Agente para buscar valores en voucher" : "Automatización de pagos"}
          </h1>
          {view === "query-memorando"
            ? <QueryMemorandoSection apiBase={API_BASE} />
            : view === "voucher"
            ? <VoucherSearchSection apiBase={API_BASE} />
            : <ProcesoUnificadoSection apiBase={API_BASE} view={view} />}
        </main>
      </div>
    </div>
  );
}

export default App;
