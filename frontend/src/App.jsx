import { useMemo, useState } from "react";
import ProcesoUnificadoSection from "./components/ProcesoUnificadoSection";

const API_BASE = (import.meta.env.VITE_API_URL || "http://localhost:8000").replace(/\/$/, "");

function decodeBase64ToBlob(base64, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: mimeType });
}

function App() {
  const [selectedFile, setSelectedFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const filename = useMemo(() => selectedFile?.name || "Sin archivo seleccionado", [selectedFile]);

  const handleFileChange = (event) => {
    const file = event.target.files?.[0] || null;
    setSelectedFile(file);
    setError("");
  };

  const handleReset = () => {
    setSelectedFile(null);
    setError("");
    setResult(null);
  };

  const triggerDownload = (base64Content, sourceName) => {
    const blob = decodeBase64ToBlob(
      base64Content,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const safeName = sourceName?.replace(/\.xlsx$/i, "") || "conciliado";
    link.href = url;
    link.download = `${safeName}_conciliado.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const handleDownloadResult = () => {
    if (result?.file) {
      triggerDownload(result.file, selectedFile?.name || "conciliado");
    }
  };

  const handleProcess = async () => {
    if (!selectedFile) {
      setError("Selecciona un archivo .xlsx antes de procesar.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const response = await fetch(`${API_BASE}/procesar`, {
        method: "POST",
        body: formData,
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload?.detail || "No se pudo procesar el archivo");
      }

      setResult(payload);
      if (payload.file) {
        triggerDownload(payload.file, selectedFile.name);
      }
    } catch (err) {
      setError(err.message || "Error inesperado durante el procesamiento");
      setResult(null);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app-shell">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;700&display=swap');

        :root {
          --bg-1: #f5f8ff;
          --bg-2: #eef4ff;
          --bg-3: #ffffff;
          --panel: rgba(255, 255, 255, 0.92);
          --panel-strong: rgba(255, 255, 255, 0.98);
          --border: rgba(80, 108, 184, 0.12);
          --border-strong: rgba(80, 108, 184, 0.18);
          --text: #14213d;
          --muted: #64748b;
          --muted-strong: #334155;
          --accent: #2563eb;
          --accent-2: #4f46e5;
          --accent-soft: rgba(37, 99, 235, 0.1);
          --success: #0f9d58;
          --warn: #d97706;
          --danger: #dc2626;
          --shadow: 0 24px 60px rgba(15, 23, 42, 0.09);
        }

        * { box-sizing: border-box; }

        body {
          margin: 0;
          font-family: "Inter", sans-serif;
          color: var(--text);
          min-height: 100vh;
          background:
            radial-gradient(circle at top left, rgba(37, 99, 235, 0.12), transparent 34%),
            radial-gradient(circle at top right, rgba(79, 70, 229, 0.08), transparent 30%),
            linear-gradient(180deg, var(--bg-1), var(--bg-2) 42%, #edf2ff 100%);
        }

        .app-shell {
          min-height: 100vh;
          padding: 24px;
          position: relative;
          overflow-x: hidden;
        }

        .app-shell::before,
        .app-shell::after {
          content: "";
          position: fixed;
          border-radius: 999px;
          filter: blur(38px);
          pointer-events: none;
          z-index: 0;
        }

        .app-shell::before {
          width: 320px;
          height: 320px;
          top: -90px;
          right: -40px;
          background: rgba(37, 99, 235, 0.12);
        }

        .app-shell::after {
          width: 260px;
          height: 260px;
          bottom: -70px;
          left: -50px;
          background: rgba(79, 70, 229, 0.08);
        }

        .app-shell > * {
          position: relative;
          z-index: 1;
        }

        .container {
          max-width: 1540px;
          margin: 0 auto;
          display: grid;
          grid-template-columns: 280px minmax(0, 1fr);
          gap: 20px;
          align-items: start;
        }

        .panel {
          background: var(--panel);
          border: 1px solid var(--border);
          border-radius: 24px;
          box-shadow: var(--shadow);
          backdrop-filter: blur(14px);
          overflow: hidden;
        }

        .panel-pad {
          padding: 22px;
        }

        .sidebar {
          position: sticky;
          top: 24px;
          display: grid;
          gap: 18px;
          align-self: start;
          padding: 20px;
          background: linear-gradient(180deg, rgba(255,255,255,0.95), rgba(248,250,255,0.9));
        }

        .brand-box {
          display: flex;
          flex-direction: column;
          align-items: flex-start;
          gap: 10px;
          padding-bottom: 16px;
          border-bottom: 1px solid rgba(37, 99, 235, 0.08);
        }

        .brand-logo {
          width: 220px;
          max-width: 100%;
          height: auto;
          display: block;
        }

        .brand-meta {
          display: grid;
          gap: 4px;
        }

        .brand-name {
          margin: 0;
          font-family: "Space Grotesk", sans-serif;
          font-size: 1.2rem;
          line-height: 1;
        }

        .brand-subtitle {
          margin: 4px 0 0;
          color: var(--muted);
          font-size: 0.84rem;
        }

        .nav-list {
          display: grid;
          gap: 8px;
        }

        .nav-item {
          display: flex;
          align-items: center;
          gap: 12px;
          padding: 12px 14px;
          border-radius: 14px;
          color: var(--muted-strong);
          background: rgba(37, 99, 235, 0.03);
          border: 1px solid transparent;
          font-weight: 600;
        }

        .nav-item.active {
          background: linear-gradient(90deg, rgba(37, 99, 235, 0.12), rgba(79, 70, 229, 0.08));
          border-color: rgba(37, 99, 235, 0.16);
          color: var(--text);
        }

        .nav-dot {
          width: 10px;
          height: 10px;
          border-radius: 999px;
          background: var(--accent);
          box-shadow: 0 0 0 6px rgba(37, 99, 235, 0.08);
          flex-shrink: 0;
        }

        .sidebar-card {
          padding: 16px;
          border-radius: 18px;
          background: linear-gradient(180deg, #fff, #f8fbff);
          border: 1px solid rgba(37, 99, 235, 0.08);
        }

        .sidebar-card h3 {
          margin: 0 0 8px;
          font-size: 0.98rem;
        }

        .sidebar-card p {
          margin: 0;
          color: var(--muted);
          line-height: 1.6;
          font-size: 0.92rem;
        }

        .toolbar {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 16px;
          padding: 18px 20px;
          margin-bottom: 18px;
          background: rgba(255,255,255,0.94);
        }

        .toolbar-title {
          display: grid;
          gap: 6px;
        }

        .toolbar-title h2 {
          margin: 0;
          font-family: "Space Grotesk", sans-serif;
          font-size: 1.35rem;
          letter-spacing: -0.03em;
        }

        .toolbar-title p {
          margin: 0;
          color: var(--muted);
          line-height: 1.5;
        }

        .toolbar-actions {
          display: flex;
          align-items: center;
          gap: 10px;
          flex-wrap: wrap;
          justify-content: flex-end;
        }

        .toolbar-badge {
          padding: 9px 12px;
          border-radius: 999px;
          border: 1px solid rgba(37, 99, 235, 0.12);
          background: rgba(37, 99, 235, 0.06);
          color: var(--muted-strong);
          font-size: 0.82rem;
          font-weight: 700;
          white-space: nowrap;
        }

        .btn {
          border: 0;
          border-radius: 14px;
          padding: 12px 18px;
          font-family: "Inter", sans-serif;
          font-weight: 700;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          transition: transform 160ms ease, box-shadow 180ms ease, filter 160ms ease, background 160ms ease;
        }

        .btn:hover:not(:disabled) {
          transform: translateY(-1px);
        }

        .btn:disabled {
          cursor: not-allowed;
          filter: grayscale(20%);
          opacity: 0.7;
        }

        .btn-primary {
          color: #fff;
          background: linear-gradient(120deg, var(--accent), var(--accent-2));
          box-shadow: 0 14px 24px rgba(37, 99, 235, 0.18);
        }

        .btn-secondary {
          color: var(--text);
          background: #fff;
          border: 1px solid rgba(37, 99, 235, 0.14);
          box-shadow: 0 10px 20px rgba(15, 23, 42, 0.04);
        }

        .btn-ghost {
          color: var(--muted-strong);
          background: rgba(255, 255, 255, 0.7);
          border: 1px solid rgba(148, 163, 184, 0.2);
        }

        .btn-warning {
          color: #1f2937;
          background: linear-gradient(120deg, #fde68a, #fef3c7);
          border: 1px solid rgba(245, 158, 11, 0.18);
        }

        .main-column {
          min-width: 0;
          display: grid;
          gap: 18px;
        }

        .hero {
          display: grid;
          grid-template-columns: 1.1fr 0.9fr;
          gap: 18px;
        }

        .brand-row {
          display: flex;
          flex-wrap: wrap;
          align-items: center;
          gap: 10px;
          margin-bottom: 16px;
        }

        .brand-pill {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 8px 12px;
          border-radius: 999px;
          border: 1px solid rgba(37, 99, 235, 0.12);
          background: rgba(37, 99, 235, 0.05);
          color: var(--muted-strong);
          font-size: 12px;
          font-weight: 700;
          letter-spacing: 0.08em;
          text-transform: uppercase;
        }

        .eyebrow, .section-eyebrow {
          color: var(--accent);
          font-weight: 800;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          font-size: 12px;
          margin: 0 0 12px;
        }

        .title {
          margin: 0;
          font-family: "Space Grotesk", sans-serif;
          font-size: clamp(2rem, 4vw, 3.5rem);
          line-height: 1;
          letter-spacing: -0.05em;
          max-width: 14ch;
          color: #101828;
        }

        .subtitle {
          margin: 14px 0 0;
          color: var(--muted);
          max-width: 62ch;
          font-size: 1.02rem;
          line-height: 1.7;
        }

        .hero-grid {
          display: grid;
          grid-template-columns: 1.1fr 0.9fr;
          gap: 18px;
          margin-top: 22px;
        }

        .upload-card {
          display: grid;
          gap: 16px;
        }

        .upload-actions {
          display: flex;
          flex-wrap: wrap;
          gap: 10px;
          align-items: center;
        }

        .input-file {
          width: 100%;
          border: 1px dashed rgba(37, 99, 235, 0.22);
          background: rgba(37, 99, 235, 0.03);
          color: var(--muted-strong);
          border-radius: 16px;
          padding: 14px;
          max-width: 420px;
        }

        .input-file::file-selector-button {
          margin-right: 14px;
          border: 0;
          border-radius: 12px;
          padding: 10px 14px;
          background: linear-gradient(120deg, var(--accent), var(--accent-2));
          color: #fff;
          font-weight: 700;
          cursor: pointer;
        }

        .filename {
          color: var(--muted);
          font-size: 0.92rem;
          min-width: 240px;
        }

        .error {
          margin: 0;
          color: var(--danger);
          font-weight: 600;
          background: rgba(220, 38, 38, 0.06);
          border: 1px solid rgba(220, 38, 38, 0.14);
          border-radius: 14px;
          padding: 12px 14px;
        }

        .meta-grid {
          display: grid;
          gap: 12px;
          align-content: start;
        }

        .stat-strip {
          display: grid;
          grid-template-columns: repeat(3, minmax(0, 1fr));
          gap: 12px;
        }

        .stat {
          padding: 16px;
          border-radius: 18px;
          background: linear-gradient(180deg, #fff, #f8fbff);
          border: 1px solid rgba(37, 99, 235, 0.08);
        }

        .stat-label {
          color: var(--muted);
          font-size: 0.82rem;
          margin: 0;
          font-weight: 600;
        }

        .stat-value {
          margin: 8px 0 0;
          font-size: 1.65rem;
          font-weight: 800;
          letter-spacing: -0.03em;
          color: #101828;
        }

        .callout {
          padding: 18px;
          border-radius: 18px;
          background: linear-gradient(180deg, rgba(37, 99, 235, 0.05), rgba(79, 70, 229, 0.03));
          border: 1px solid rgba(37, 99, 235, 0.12);
          color: var(--muted-strong);
        }

        .callout h3 {
          margin: 0 0 8px;
          font-size: 1rem;
        }

        .callout p {
          margin: 0;
          line-height: 1.6;
          color: var(--muted);
        }

        .feature-list {
          margin: 0;
          padding-left: 18px;
          display: grid;
          gap: 10px;
          color: var(--muted-strong);
          line-height: 1.45;
        }

        .feature-list li::marker {
          color: var(--accent);
        }

        .dashboard-shell {
          margin-top: 2px;
        }

        .summary-grid {
          display: grid;
          grid-template-columns: repeat(3, minmax(0, 1fr));
          gap: 14px;
        }

        .summary-card {
          position: relative;
          overflow: hidden;
          padding: 18px 18px 20px;
          border-radius: 22px;
          background: linear-gradient(180deg, #fff, #f8fbff);
          border: 1px solid rgba(37, 99, 235, 0.08);
          box-shadow: 0 16px 42px rgba(15, 23, 42, 0.05);
        }

        .summary-card::before {
          content: "";
          position: absolute;
          inset: 0 auto 0 0;
          width: 4px;
          background: var(--tone);
        }

        .summary-accent {
          position: absolute;
          top: -42px;
          right: -18px;
          width: 120px;
          height: 120px;
          border-radius: 999px;
          background: radial-gradient(circle, color-mix(in srgb, var(--tone) 28%, transparent), transparent 70%);
          opacity: 0.9;
        }

        .summary-label {
          margin: 0;
          color: var(--muted);
          font-size: 0.84rem;
          font-weight: 600;
        }

        .summary-value {
          margin: 10px 0 0;
          font-family: "Space Grotesk", sans-serif;
          font-size: clamp(1.8rem, 3vw, 2.5rem);
          letter-spacing: -0.04em;
          color: #101828;
        }

        .summary-subtitle {
          margin: 8px 0 0;
          color: var(--muted);
          font-size: 0.85rem;
          font-weight: 500;
        }

        .alerts-stack {
          display: grid;
          gap: 10px;
        }

        .alert-card {
          display: flex;
          gap: 10px;
          align-items: flex-start;
          padding: 14px 16px;
          border-radius: 18px;
          background: linear-gradient(180deg, rgba(245, 158, 11, 0.08), rgba(245, 158, 11, 0.04));
          border: 1px solid rgba(245, 158, 11, 0.18);
          color: #7c4d00;
        }

        .alert-card strong {
          color: #a16207;
          min-width: 70px;
        }

        .logs-panel {
          border-radius: 24px;
          overflow: hidden;
          background: linear-gradient(180deg, #fff, #fbfdff);
          border: 1px solid rgba(37, 99, 235, 0.08);
          box-shadow: 0 24px 70px rgba(15, 23, 42, 0.06);
        }

        .logs-header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 16px;
          padding: 20px 22px 16px;
          border-bottom: 1px solid rgba(37, 99, 235, 0.08);
        }

        .logs-header h2 {
          margin: 0;
          font-family: "Space Grotesk", sans-serif;
          font-size: 1.3rem;
          color: #101828;
        }

        .logs-count {
          padding: 9px 12px;
          border-radius: 999px;
          background: rgba(37, 99, 235, 0.06);
          color: var(--muted-strong);
          font-size: 0.82rem;
          font-weight: 700;
          white-space: nowrap;
        }

        .logs-empty {
          margin: 0;
          padding: 20px 22px 22px;
          color: var(--muted);
        }

        .logs-table {
          width: 100%;
          border-collapse: collapse;
          min-width: 980px;
        }

        .logs-table thead th {
          padding: 14px 18px;
          text-align: left;
          font-size: 0.8rem;
          text-transform: uppercase;
          letter-spacing: 0.08em;
          color: var(--muted);
          background: #f8fbff;
        }

        .logs-table tbody td {
          padding: 16px 18px;
          border-top: 1px solid rgba(37, 99, 235, 0.08);
          color: var(--muted-strong);
          vertical-align: top;
          background: #fff;
        }

        .logs-table tbody tr:nth-child(even) td {
          background: #fbfdff;
        }

        .logs-value {
          font-weight: 700;
          color: #101828 !important;
          white-space: nowrap;
        }

        .logs-detail {
          min-width: 420px;
          line-height: 1.6;
          color: var(--muted);
        }

        .empty-state {
          padding: 18px 22px 22px;
          border-radius: 24px;
          background: linear-gradient(180deg, #fff, #f8fbff);
          border: 1px solid rgba(37, 99, 235, 0.08);
          color: var(--muted);
        }

        .empty-state h3 {
          margin: 0 0 8px;
          color: #101828;
          font-family: "Space Grotesk", sans-serif;
        }

        .dashboard-grid {
          display: grid;
          gap: 14px;
        }

        .content-grid {
          display: grid;
          gap: 18px;
        }

        @media (max-width: 1100px) {
          .container {
            grid-template-columns: 1fr;
          }

          .sidebar {
            position: relative;
            top: 0;
          }
        }

        @media (max-width: 768px) {
          .app-shell { padding: 14px; }
          .hero,
          .hero-grid,
          .stat-strip,
          .summary-grid {
            grid-template-columns: 1fr;
          }

          .toolbar {
            flex-direction: column;
            align-items: flex-start;
          }

          .title {
            max-width: none;
          }

          .input-file,
          .filename {
            max-width: none;
            width: 100%;
          }

          .upload-actions {
            align-items: stretch;
          }

          .logs-header {
            flex-direction: column;
            align-items: flex-start;
          }

          .logs-detail {
            min-width: 280px;
          }
        }
      `}</style>

      <main className="container">
        <aside className="panel sidebar">
          <div className="brand-box">
            <img className="brand-logo" src="/calypso-logo.svg" alt="CALYPSO" />
            <div className="brand-meta">
              <h1 className="brand-name">Conciliador 2026</h1>
              <p className="brand-subtitle">Automatización de cruces contables</p>
            </div>
          </div>

          <nav className="nav-list" aria-label="Navegación principal">
            <div className="nav-item active"><span className="nav-dot" />Dashboard</div>
          </nav>

          <div className="sidebar-card">
            <h3>Estado del sistema</h3>
            <p>
              Procesamiento en memoria, sin base de datos y listo para revisar logs, alertas y resultados.
            </p>
          </div>

        </aside>

        <section className="main-column">
          <ProcesoUnificadoSection apiBase={API_BASE} />
        </section>
      </main>
    </div>
  );
}

export default App;
