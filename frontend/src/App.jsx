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
      <path d="M3.5 10.5 12 4l8.5 6.5" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M5.5 9.5v9h13v-9M9.5 18.5v-5h5v5" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function IconProcesos() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <rect x="3" y="3" width="7" height="7" rx="2" stroke="currentColor" strokeWidth="1.7" />
      <rect x="14" y="3" width="7" height="7" rx="2" stroke="currentColor" strokeWidth="1.7" />
      <rect x="3" y="14" width="7" height="7" rx="2" stroke="currentColor" strokeWidth="1.7" />
      <rect x="14" y="14" width="7" height="7" rx="2" stroke="currentColor" strokeWidth="1.7" />
      <path d="M10 6.5h4M6.5 10v4M17.5 10v4M10 17.5h4" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconConfig() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="1.7" />
      <path d="M12 2.8v2M12 19.2v2M21.2 12h-2M4.8 12h-2M18.5 5.5l-1.4 1.4M6.9 17.1l-1.4 1.4M18.5 18.5l-1.4-1.4M6.9 6.9 5.5 5.5" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconQuery() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M6 3h8l4 4v14H6a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M14 3v5h5M8 11h7M8 15h4" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
      <circle cx="16.5" cy="16.5" r="2.5" stroke="currentColor" strokeWidth="1.7" />
      <path d="m18.3 18.3 2 2" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconVoucher() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M6 3h12v18H6z" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" />
      <path d="M9 7h6M9 11h2M13 11h2M9 15h6M9 18h4" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
      <circle cx="16.8" cy="16.8" r="3.2" fill="var(--card)" stroke="currentColor" strokeWidth="1.7" />
      <path d="m15.8 16.8 0.8 0.8 1.6-1.7" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function IconContabilidad() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <rect x="3" y="4" width="18" height="16" rx="2" stroke="currentColor" strokeWidth="1.7" />
      <path d="M7 8h10M7 12h3M14 12h3M7 16h3M14 16h3" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function StatusIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.7" />
      <path d="m7.8 12.2 2.7 2.7 5.7-6" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round" />
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
            <button className={`nav-item ${view === "contabilidad-cruces" ? "active" : ""}`} onClick={() => setView("contabilidad-cruces")}>
              <IconContabilidad />
              Contabilidad cruces
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
            <span className="status-dot" aria-hidden="true"><StatusIcon /></span>
            <div>
              <strong>Sistema actualizado</strong>
              <span>v{APP_VERSION}</span>
            </div>
          </div>
        </aside>

        <main className="workspace">
          <h1 className="page-title">
            {view === "contabilidad-cruces"
              ? "Contabilidad cruces"
              : view === "query-memorando"
              ? "Conciliar Query y memorando ya conciliado"
              : view === "voucher"
              ? "Agente para buscar valores en voucher"
              : "Automatización de pagos"}
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
