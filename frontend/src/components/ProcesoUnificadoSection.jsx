import { useState } from "react";
import Dashboard from "./Dashboard";
import { APP_VERSION } from "../generated/buildInfo";

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

function CloudIcon() {
  return (
    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M7 18h10a4 4 0 0 0 .6-8 6 6 0 0 0-11.4-1.5A4.5 4.5 0 0 0 7 18Z" stroke="currentColor" strokeWidth="1.8" />
      <path d="M12 11v6M12 11l-2.2 2.2M12 11l2.2 2.2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    </svg>
  );
}

function FileDropCard({ title, hint, file, onFile, disabled, inputKey }) {
  const [dragover, setDragover] = useState(false);

  const assignFile = (nextFile) => {
    if (nextFile) onFile(nextFile);
  };

  return (
    <label className={`drop-card ${file ? "has-file" : ""} ${dragover ? "dragover" : ""}`}>
      <input
        key={inputKey}
        type="file"
        accept=".xlsx"
        aria-label={title}
        disabled={disabled}
        onChange={(event) => assignFile(event.target.files?.[0] || null)}
        onDragOver={(event) => {
          event.preventDefault();
          setDragover(true);
        }}
        onDragLeave={() => setDragover(false)}
        onDrop={(event) => {
          event.preventDefault();
          setDragover(false);
          assignFile(event.dataTransfer.files?.[0] || null);
        }}
      />
      <span className="drop-icon">
        <CloudIcon />
      </span>
      <strong>{title}</strong>
      <span className="hint">{hint}</span>
      <span className="underline" />
      {file ? <span className="filename">{file.name}</span> : null}
    </label>
  );
}

function ProcesoUnificadoSection({ apiBase, view = "inicio" }) {
  const [pseFile, setPseFile] = useState(null);
  const [memoFile, setMemoFile] = useState(null);
  const [queryInternoFile, setQueryInternoFile] = useState(null);
  const [adquirenciasFile, setAdquirenciasFile] = useState(null);
  const [dateTolerance, setDateTolerance] = useState(1);
  const [valueTolerance, setValueTolerance] = useState(0.01);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [resetToken, setResetToken] = useState(0);

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
    setResetToken((value) => value + 1);
  };

  const contableResult = result?.contable || (result?.resumen && result?.logs ? result : null);
  const pseResult = result?.pse || (result?.dataset ? result : null);

  if (view === "configuracion") {
    return (
      <section className="config-panel">
        <div className="panel-heading">
          <span className="cloud-badge"><CloudIcon /></span>
          <h2>Configuración</h2>
        </div>
        <p>
          La conciliación sigue usando el mismo backend y las mismas reglas de cruce. Aquí solo se muestra
          el entorno activo, sin cambiar tolerancias ni archivos.
        </p>
        <p><strong>API:</strong> {apiBase}</p>
        <p><strong>Versión:</strong> v{APP_VERSION}</p>
      </section>
    );
  }

  if (view === "procesos") {
    return (
      <section className="results-panel" style={{ display: "grid", gap: 16 }}>
        <div className="panel-heading">
          <span className="cloud-badge"><CloudIcon /></span>
          <h2>Procesos</h2>
        </div>
        {!result && !loading ? (
          <div className="empty-state">
            <h3>Sin procesos recientes</h3>
            <p>Carga los archivos en Inicio y pulsa Procesar archivos para ver el resumen y la trazabilidad aquí.</p>
          </div>
        ) : (
          <>
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
          </>
        )}
      </section>
    );
  }

  return (
    <section className="upload-panel">
      <div className="panel-heading">
        <span className="cloud-badge"><CloudIcon /></span>
        <h2>Carga de archivos</h2>
      </div>

      <div className="drop-grid">
        <FileDropCard
          title="PSE"
          hint="Archivo requerido"
          file={pseFile}
          onFile={setPseFile}
          disabled={loading}
          inputKey={`pse-${resetToken}`}
        />
        <FileDropCard
          title="Conciliación"
          hint="Archivo requerido"
          file={memoFile}
          onFile={setMemoFile}
          disabled={loading}
          inputKey={`memo-${resetToken}`}
        />
        <FileDropCard
          title="Query Interno"
          hint="Opcional"
          file={queryInternoFile}
          onFile={setQueryInternoFile}
          disabled={loading}
          inputKey={`query-${resetToken}`}
        />
        <FileDropCard
          title="Adquirencias"
          hint="Opcional"
          file={adquirenciasFile}
          onFile={setAdquirenciasFile}
          disabled={loading}
          inputKey={`adq-${resetToken}`}
        />
      </div>

      <div className="refs-grid">
        <div className="ref-field">
          <label>
            <span className="info-icon" title="Tolerancia en días para PSE">i</span>
            Valor de referencia
          </label>
          <input
            type="number"
            min="0"
            step="1"
            value={dateTolerance}
            onChange={(event) => setDateTolerance(Number(event.target.value || 0))}
            disabled={loading}
            aria-label="Tolerancia en días"
          />
        </div>
        <div className="ref-field">
          <label>
            <span className="info-icon">i</span>
            Valor de referencia
          </label>
          <input type="text" disabled aria-hidden="true" />
        </div>
        <div className="ref-field">
          <label>
            <span className="info-icon" title="Tolerancia en valor">i</span>
            Valor de referencia
          </label>
          <input
            type="number"
            min="0"
            step="0.01"
            value={valueTolerance}
            onChange={(event) => setValueTolerance(Number(event.target.value || 0))}
            disabled={loading}
            aria-label="Tolerancia en valor"
          />
        </div>
        <div className="ref-field">
          <label>
            <span className="info-icon">i</span>
            Valor de referencia
          </label>
          <input type="text" disabled aria-hidden="true" />
        </div>
      </div>

      <div className="action-row">
        <button className="btn btn-primary" onClick={handleProcess} disabled={loading}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
            <path d="M8 5v14l11-7z" />
          </svg>
          {loading ? "Procesando..." : "Procesar archivos"}
        </button>
        <button className="btn btn-secondary" onClick={() => downloadResult(result)} disabled={!result}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
            <path d="M12 4v10" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
            <path d="M8 10l4 4 4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M5 19h14" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
          </svg>
          Descargar resultados
        </button>
        <span className="spacer" />
        <button className="btn btn-ghost" onClick={handleReset} disabled={loading}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
            <path d="M4 7h10a6 6 0 1 1 0 12H9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
            <path d="M8 3 4 7l4 4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
          Limpiar
        </button>
      </div>

      {error && <p className="error">{error}</p>}
    </section>
  );
}

export default ProcesoUnificadoSection;
