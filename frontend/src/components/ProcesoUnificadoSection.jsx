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
    border: hasFile ? "1px solid #22c55e" : "1px solid rgba(37, 99, 235, 0.14)",
    background: hasFile ? "rgba(34, 197, 94, 0.12)" : "#fff",
    borderRadius: 16,
    padding: "12px 14px",
    color: "var(--text)",
    fontSize: 14,
    boxShadow: hasFile ? "0 0 0 3px rgba(34, 197, 94, 0.12)" : "none",
    transition: "all 0.2s ease",
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
        <span className="brand-pill">Proceso unificado</span>
        <span style={{ color: "var(--muted)", fontSize: 13 }}>
          Un solo envío para conciliación PSE y memorando de cruces
        </span>
      </div>

      <div>
        <p className="eyebrow">Flujo único</p>
        <h2 className="title" style={{ maxWidth: "18ch", fontSize: "clamp(1.7rem, 3vw, 2.6rem)" }}>
          Dos archivos, un solo cruce
        </h2>
        <p className="subtitle" style={{ maxWidth: "74ch" }}>
          Carga el archivo PSE y el memorando de conciliación contable. El backend cruza los valores exactos entre
          ambos archivos y devuelve el resultado conciliado sin alterar el formato original.
        </p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 12 }}>
        <div style={{ display: "grid", gap: 8 }}>
          <input
            className="input-file"
            type="file"
            accept=".xlsx"
            aria-label="Archivo PSE"
            onChange={(event) => setPseFile(event.target.files?.[0] || null)}
            disabled={loading}
            style={fileInputStyle(Boolean(pseFile))}
          />
          <span style={{ color: "var(--muted)", fontSize: 12, lineHeight: 1.4 }}>
            Sube aquí el archivo PSE.
          </span>
        </div>
        <div style={{ display: "grid", gap: 8 }}>
          <input
            className="input-file"
            type="file"
            accept=".xlsx"
            aria-label="Archivo de conciliación contable"
            onChange={(event) => setMemoFile(event.target.files?.[0] || null)}
            disabled={loading}
            style={fileInputStyle(Boolean(memoFile))}
          />
          <span style={{ color: "var(--muted)", fontSize: 12, lineHeight: 1.4 }}>
            Sube aquí el archivo de conciliación contable o memorando de cruces.
          </span>
        </div>
        <div style={{ display: "grid", gap: 8 }}>
          <input
            className="input-file"
            type="file"
            accept=".xlsx"
            aria-label="Archivo de Query Interno"
            onChange={(event) => setQueryInternoFile(event.target.files?.[0] || null)}
            disabled={loading}
            style={fileInputStyle(Boolean(queryInternoFile))}
          />
          <span style={{ color: "var(--muted)", fontSize: 12, lineHeight: 1.4 }}>
            Sube aquí el archivo de Query Interno (opcional) - Segundo cruce interno.
          </span>
        </div>
        <div style={{ display: "grid", gap: 8 }}>
          <input
            className="input-file"
            type="file"
            accept=".xlsx"
            aria-label="Archivo de Adquirencias"
            onChange={(event) => setAdquirenciasFile(event.target.files?.[0] || null)}
            disabled={loading}
            style={fileInputStyle(Boolean(adquirenciasFile))}
          />
          <span style={{ color: "var(--muted)", fontSize: 12, lineHeight: 1.4 }}>
            Sube aquí el archivo de Adquirencias (opcional).
          </span>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12 }}>
        <input
          type="number"
          min="0"
          step="1"
          value={dateTolerance}
          onChange={(event) => setDateTolerance(Number(event.target.value || 0))}
          placeholder="Tolerancia en días"
          style={{ width: "100%", border: "1px solid rgba(37, 99, 235, 0.14)", background: "#fff", borderRadius: 16, padding: 14, color: "var(--text)", fontSize: 14 }}
          disabled={loading}
        />
        <input
          type="number"
          min="0"
          step="0.01"
          value={valueTolerance}
          onChange={(event) => setValueTolerance(Number(event.target.value || 0))}
          placeholder="Tolerancia en valor"
          style={{ width: "100%", border: "1px solid rgba(37, 99, 235, 0.14)", background: "#fff", borderRadius: 16, padding: 14, color: "var(--text)", fontSize: 14 }}
          disabled={loading}
        />
      </div>

      <div className="upload-actions">
        <button className="btn btn-primary" onClick={handleProcess} disabled={loading}>
          {loading ? "Procesando..." : "Procesar todo"}
        </button>
        <button className="btn btn-secondary" onClick={() => downloadResult(result)} disabled={!result}>
          Descargar resultados
        </button>
        <button className="btn btn-ghost" onClick={handleReset} disabled={loading}>
          Limpiar
        </button>
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
