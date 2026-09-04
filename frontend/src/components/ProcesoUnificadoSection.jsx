import { useState } from "react";
import Dashboard from "./Dashboard";

function decodeBase64ToBlob(base64, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: mimeType });
}

function downloadBase64(fileContent, name) {
  const blob = decodeBase64ToBlob(
    fileContent,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function ProcesoUnificadoSection({ apiBase, apiKey }) {
  const [pseFile, setPseFile] = useState(null);
  const [memoFile, setMemoFile] = useState(null);
  const [queryInternoFile, setQueryInternoFile] = useState(null);
  const [adquirenciasFile, setAdquirenciasFile] = useState(null);
  const [dateTolerance, setDateTolerance] = useState(1);
  const [valueTolerance, setValueTolerance] = useState(0.01);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const downloadResult = (payload) => {
    if (Array.isArray(payload?.files) && payload.files.length > 0) {
      payload.files.forEach((item) => {
        if (item?.file && item?.name) {
          downloadBase64(item.file, item.name);
        }
      });
      return;
    }

    if (payload?.file) {
      downloadBase64(payload.file, payload.output_name || "resultado.xlsx");
    }
    if (payload?.secondary_file) {
      downloadBase64(payload.secondary_file, payload.secondary_output_name || "resultado_secundario.xlsx");
    }
  };

  const fileInputStyle = (hasFile) => ({
    width: "100%",
    minHeight: 56,
    border: hasFile ? "1px solid #22c55e" : "1px solid rgba(37, 99, 235, 0.14)",
    background: hasFile ? "rgba(34, 197, 94, 0.12)" : "#fff",
    borderRadius: 16,
    padding: "12px 14px",
    color: "var(--text)",
    fontSize: 14,
    boxShadow: hasFile ? "0 0 0 3px rgba(34, 197, 94, 0.12)" : "none",
    transition: "all 0.2s ease",
    display: "block",
    boxSizing: "border-box",
  });

  const handleProcess = async () => {
    if (!memoFile) {
      setError("Selecciona al menos el archivo de conciliación contable.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const formData = new FormData();
      if (pseFile) {
        formData.append("pse_file", pseFile);
      }
      if (queryInternoFile) {
        formData.append("query_interno_file", queryInternoFile);
      }
      if (adquirenciasFile) {
        formData.append("adquirencias_file", adquirenciasFile);
      }
      if (memoFile) {
        formData.append("file", memoFile);
        formData.append("cruces_file", memoFile);
      }
      formData.append("tolerance_days", String(dateTolerance));
      formData.append("tolerance_value", String(valueTolerance));

      const response = await fetch(`${apiBase}/procesar`, {
        method: "POST",
        headers: apiKey ? { "X-API-Key": apiKey } : undefined,
        body: formData,
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload?.detail || "No se pudo procesar la información");
      }

      setResult(payload);
      downloadResult(payload);
    } catch (err) {
      setError(err.message || "Error inesperado durante el procesamiento");
      setResult(null);
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
    setPseFile(null);
    setMemoFile(null);
    setQueryInternoFile(null);
    setAdquirenciasFile(null);
    setError("");
    setResult(null);
  };

  const contableResult = result?.contable || (result?.resumen && result?.logs ? result : null);
  const pseResult = result?.pse || (result?.dataset ? result : null);

  return (
    <section className="panel panel-pad" style={{ display: "grid", gap: 16 }}>
      <div className="brand-row" style={{ marginBottom: 0 }}>
        <span className="brand-pill">Automatización de pagos</span>
      </div>

      <div>
        <p className="eyebrow">Carga de archivos</p>
        <h2 className="title" style={{ maxWidth: "48ch", fontSize: "clamp(1.7rem, 3vw, 2.6rem)" }}>
          Automatización de pagos
        </h2>
        <p className="subtitle" style={{ maxWidth: "74ch" }}>
          Carga los archivos necesarios para el proceso de conciliación y consulta de cruces.
        </p>
      </div>

      <div className="upload-card">
        <div className="upload-header">
          <div style={{ width: 44, height: 44, borderRadius: 10, background: 'linear-gradient(180deg, rgba(37,99,235,0.08), rgba(79,70,229,0.04))', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent)' }}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
              <path d="M12 3v8" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
              <path d="M8 7l4-4 4 4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </div>
          <div>
            <div style={{ fontWeight: 700 }}>Carga de archivos</div>
            <div style={{ fontSize: 12, color: 'var(--muted)' }}>Selecciona los archivos para iniciar el proceso</div>
          </div>
        </div>

        <div className="upload-grid">
          {/* PSE */}
          <div className="upload-tile-wrapper">
            <div className={`upload-tile ${pseFile ? 'loaded' : ''}`}>
              <div className="upload-tile-inner">
                <div className="upload-icon">
                  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
                    <path d="M8 12l4-4 4 4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
                    <path d="M12 16V8" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
                  </svg>
                </div>
                <div className="upload-label">PSE</div>
                <div className="upload-sub">Archivo requerido</div>
                {pseFile && <div className="upload-filename">{pseFile.name}</div>}
                <input className="upload-input" type="file" accept=".xlsx" aria-label="Archivo PSE" onChange={(event) => setPseFile(event.target.files?.[0] || null)} disabled={loading} />
              </div>
              <div className="ref-row">
                <div style={{ width: 18, height: 18, borderRadius: 6, background: 'rgba(37,99,235,0.06)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent)' }}>
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
                    <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="1.2" />
                    <path d="M12 8v4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                    <circle cx="12" cy="16.4" r="0.8" fill="currentColor" />
                  </svg>
                </div>
                <div style={{ color: 'var(--muted)', fontSize: 13, minWidth: 120 }}>Valor de referencia</div>
                <input className="small-input" type="number" min="0" step="1" value={dateTolerance} onChange={(event) => setDateTolerance(Number(event.target.value || 0))} disabled={loading} />
              </div>
            </div>
          </div>

          {/* Conciliación */}
          <div className="upload-tile-wrapper">
            <div className={`upload-tile ${memoFile ? 'loaded' : ''}`}>
              <div className="upload-tile-inner">
                <div className="upload-icon">
                  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
                    <path d="M8 12l4-4 4 4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
                    <path d="M12 16V8" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
                  </svg>
                </div>
                <div className="upload-label">Conciliación</div>
                <div className="upload-sub">Archivo requerido</div>
                {memoFile && <div className="upload-filename">{memoFile.name}</div>}
                <input className="upload-input" type="file" accept=".xlsx" aria-label="Archivo de conciliación contable" onChange={(event) => setMemoFile(event.target.files?.[0] || null)} disabled={loading} />
              </div>
              <div className="ref-row">
                <div style={{ width: 18, height: 18, borderRadius: 6, background: 'rgba(37,99,235,0.06)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent)' }}>
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
                    <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="1.2" />
                    <path d="M12 8v4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                    <circle cx="12" cy="16.4" r="0.8" fill="currentColor" />
                  </svg>
                </div>
                <div style={{ color: 'var(--muted)', fontSize: 13, minWidth: 120 }}>Valor de referencia</div>
                <input className="small-input" type="text" value={""} onChange={() => {}} disabled={loading} />
              </div>
            </div>
          </div>

          {/* Query Interno */}
          <div className="upload-tile-wrapper">
            <div className={`upload-tile ${queryInternoFile ? 'loaded' : ''}`}>
              <div className="upload-tile-inner">
                <div className="upload-icon">
                  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
                    <path d="M8 12l4-4 4 4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
                    <path d="M12 16V8" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
                  </svg>
                </div>
                <div className="upload-label">Query Interno</div>
                <div className="upload-sub">Opcional</div>
                {queryInternoFile && <div className="upload-filename">{queryInternoFile.name}</div>}
                <input className="upload-input" type="file" accept=".xlsx" aria-label="Archivo de Query Interno" onChange={(event) => setQueryInternoFile(event.target.files?.[0] || null)} disabled={loading} />
              </div>
              <div className="ref-row">
                <div style={{ width: 18, height: 18, borderRadius: 6, background: 'rgba(37,99,235,0.06)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent)' }}>
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
                    <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="1.2" />
                    <path d="M12 8v4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                    <circle cx="12" cy="16.4" r="0.8" fill="currentColor" />
                  </svg>
                </div>
                <div style={{ color: 'var(--muted)', fontSize: 13, minWidth: 120 }}>Valor de referencia</div>
                <input className="small-input" type="number" min="0" step="0.01" value={valueTolerance} onChange={(event) => setValueTolerance(Number(event.target.value || 0))} disabled={loading} />
              </div>
            </div>
          </div>

          {/* Adquirencias */}
          <div className="upload-tile-wrapper">
            <div className={`upload-tile ${adquirenciasFile ? 'loaded' : ''}`}>
              <div className="upload-tile-inner">
                <div className="upload-icon">
                  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
                    <path d="M8 12l4-4 4 4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
                    <path d="M12 16V8" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
                  </svg>
                </div>
                <div className="upload-label">Adquirencias</div>
                <div className="upload-sub">Opcional</div>
                {adquirenciasFile && <div className="upload-filename">{adquirenciasFile.name}</div>}
                <input className="upload-input" type="file" accept=".xlsx" aria-label="Archivo de Adquirencias" onChange={(event) => setAdquirenciasFile(event.target.files?.[0] || null)} disabled={loading} />
              </div>
              <div className="ref-row">
                <div style={{ width: 18, height: 18, borderRadius: 6, background: 'rgba(37,99,235,0.06)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent)' }}>
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
                    <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="1.2" />
                    <path d="M12 8v4" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                    <circle cx="12" cy="16.4" r="0.8" fill="currentColor" />
                  </svg>
                </div>
                <div style={{ color: 'var(--muted)', fontSize: 13, minWidth: 120 }}>Valor de referencia</div>
                <input className="small-input" type="text" value={""} onChange={() => {}} disabled={loading} />
              </div>
            </div>
          </div>
        </div>

        <div className="actions-row">
          <div className="actions-left">
            <button className="btn-primary" onClick={handleProcess} disabled={loading}>{loading ? 'Procesando...' : 'Procesar archivos'}</button>
            <button className="btn-secondary" onClick={() => downloadResult(result)} disabled={!result}>Descargar resultados</button>
          </div>
          <div>
            <button className="btn-ghost-right" onClick={handleReset} disabled={loading}>Limpiar</button>
          </div>
        </div>
      </div>

      {error && <p className="error">{error}</p>}

      {(loading || result || pseFile || memoFile) && (
        <div style={{ display: "grid", gap: 14 }}>
          {contableResult && <Dashboard data={contableResult} loading={loading} />}

          {pseResult?.dataset && (
            <div className="logs-panel">
              <header className="logs-header">
                <div>
                  <p className="section-eyebrow">PSE conciliado</p>
                  <h2>Resultado del archivo PSE</h2>
                </div>
                <span className="logs-count">{pseResult.dataset.length} filas</span>
              </header>

              <div style={{ overflowX: "auto" }}>
                <table className="logs-table">
                  <thead>
                    <tr>
                      <th>Estado</th>
                      <th>Cuenta contable</th>
                      <th>Valores asociados</th>
                      <th>Comentario</th>
                      <th>Grupo</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pseResult.dataset.map((item) => (
                      <tr key={`${item.sheet}-${item.row}`}>
                        <td>{item.estado_conciliacion}</td>
                        <td>{item.cuenta_contable || "-"}</td>
                        <td className="logs-detail">{item.valores_asociados || "-"}</td>
                        <td className="logs-detail">{item.comentario_conciliacion}</td>
                        <td>{item.id_grupo_conciliacion || "-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {pseResult?.dataset_cruces && pseResult.dataset_cruces.length > 0 && (
            <div className="logs-panel">
              <header className="logs-header">
                <div>
                  <p className="section-eyebrow">Documento de cruces</p>
                  <h2>Detalle conciliado del archivo bancario</h2>
                </div>
                <span className="logs-count">{pseResult.dataset_cruces.length} filas</span>
              </header>

              <div style={{ overflowX: "auto" }}>
                <table className="logs-table">
                  <thead>
                    <tr>
                      <th>Estado</th>
                      <th>Cuenta contable</th>
                      <th>Valores asociados</th>
                      <th>Comentario</th>
                      <th>Grupo</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pseResult.dataset_cruces.map((item) => (
                      <tr key={`${item.sheet}-${item.row}`}>
                        <td>{item.estado_conciliacion}</td>
                        <td>{item.cuenta_contable || "-"}</td>
                        <td className="logs-detail">{item.valores_asociados || "-"}</td>
                        <td className="logs-detail">{item.comentario_conciliacion}</td>
                        <td>{item.id_grupo_conciliacion || "-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}
    </section>
  );
}

export default ProcesoUnificadoSection;
