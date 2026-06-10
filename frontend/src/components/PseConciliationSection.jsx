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

function PseConciliationSection({ apiBase }) {
  const [pseFile, setPseFile] = useState(null);
  const [crucesFile, setCrucesFile] = useState(null);
  const [dateTolerance, setDateTolerance] = useState(1);
  const [valueTolerance, setValueTolerance] = useState(0.01);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const downloadPayload = (payload) => {
    const downloads = [
      payload?.file && { file: payload.file, name: payload.output_name || "PSE_CONCILIADO.xlsx" },
      payload?.secondary_file && {
        file: payload.secondary_file,
        name: payload.secondary_output_name || "CRUCES_CONCILIADOS.xlsx",
      },
    ].filter(Boolean);

    downloads.forEach(({ file, name }) => {
      const blob = decodeBase64ToBlob(
        file,
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
    });
  };

  const handleDownload = () => {
    downloadPayload(result);
  };

  const handleProcess = async () => {
    if (!pseFile || !crucesFile) {
      setError("Selecciona el archivo PSE y el archivo de cruces contables.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const formData = new FormData();
      formData.append("pse_file", pseFile);
      formData.append("cruces_file", crucesFile);
      formData.append("tolerance_days", String(dateTolerance));
      formData.append("tolerance_value", String(valueTolerance));

      const response = await fetch(`${apiBase}/pse/conciliar`, {
        method: "POST",
        body: formData,
        credentials: "include",
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload?.detail || "No se pudo conciliar el PSE");
      }

      setResult(payload);
      if (payload.file) {
        downloadPayload(payload);
      }
    } catch (err) {
      setError(err.message || "Error inesperado durante la conciliación PSE");
      setResult(null);
    } finally {
      setLoading(false);
    }
  };

  const dataset = result?.dataset || [];
  const datasetCruces = result?.dataset_cruces || [];
  const resumen = result?.resumen || { cruzados: 0, posibles: 0, precision_estimada: 0 };

  return (
    <section className="panel panel-pad" style={{ display: "grid", gap: 16 }}>
      <div className="brand-row" style={{ marginBottom: 0 }}>
        <span className="brand-pill">Conciliación PSE</span>
        <span style={{ color: "var(--muted)", fontSize: 13 }}>
          Modo opcional sin alterar el archivo original
        </span>
      </div>

      <div>
        <p className="eyebrow">Extensión modular</p>
        <h2 className="title" style={{ maxWidth: "18ch", fontSize: "clamp(1.7rem, 3vw, 2.6rem)" }}>
          PSE conciliado con archivo adicional y columnas al final
        </h2>
        <p className="subtitle" style={{ maxWidth: "70ch" }}>
          Carga el PSE y el archivo de cruces contables. El sistema conserva la estructura base del PSE y genera
          un archivo conciliado con Estado_Conciliacion, Cuenta_Contable, Valores_Asociados,
          Comentario_Conciliacion e ID_Grupo_Conciliacion.
        </p>
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))",
          gap: 12,
        }}
      >
        <input
          className="input-file"
          type="file"
          accept=".xlsx"
          onChange={(event) => setPseFile(event.target.files?.[0] || null)}
          disabled={loading}
        />
        <input
          className="input-file"
          type="file"
          accept=".xlsx"
          onChange={(event) => setCrucesFile(event.target.files?.[0] || null)}
          disabled={loading}
        />
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
          gap: 12,
        }}
      >
        <input
          type="number"
          min="0"
          step="1"
          value={dateTolerance}
          onChange={(event) => setDateTolerance(Number(event.target.value || 0))}
          placeholder="Tolerancia en días"
          style={{
            width: "100%",
            border: "1px solid rgba(37, 99, 235, 0.14)",
            background: "#fff",
            borderRadius: 16,
            padding: 14,
            color: "var(--text)",
            fontSize: 14,
          }}
          disabled={loading}
        />
        <input
          type="number"
          min="0"
          step="0.01"
          value={valueTolerance}
          onChange={(event) => setValueTolerance(Number(event.target.value || 0))}
          placeholder="Tolerancia en valor"
          style={{
            width: "100%",
            border: "1px solid rgba(37, 99, 235, 0.14)",
            background: "#fff",
            borderRadius: 16,
            padding: 14,
            color: "var(--text)",
            fontSize: 14,
          }}
          disabled={loading}
        />
      </div>

      <div className="upload-actions">
        <button className="btn btn-primary" onClick={handleProcess} disabled={loading}>
          {loading ? "Conciliando..." : "Conciliar PSE"}
        </button>
        <button className="btn btn-secondary" onClick={handleDownload} disabled={!result?.file}>
          Descargar ambos archivos
        </button>
        <button
          className="btn btn-ghost"
          onClick={() => {
            setPseFile(null);
            setCrucesFile(null);
            setError("");
            setResult(null);
          }}
          disabled={loading}
        >
          Limpiar
        </button>
      </div>

      {error && <p className="error">{error}</p>}

      {(loading || result || pseFile || crucesFile) && <Dashboard data={result} loading={loading} />}

      {dataset.length > 0 && (
        <div className="logs-panel">
          <header className="logs-header">
            <div>
              <p className="section-eyebrow">Dataset de conciliación</p>
              <h2>Detalle PSE enriquecido</h2>
            </div>
            <span className="logs-count">{dataset.length} filas</span>
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
                {dataset.map((item) => (
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

      {datasetCruces.length > 0 && (
        <div className="logs-panel">
          <header className="logs-header">
            <div>
              <p className="section-eyebrow">Documento de cruces</p>
              <h2>Detalle conciliado del archivo bancario</h2>
            </div>
            <span className="logs-count">{datasetCruces.length} filas</span>
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
                {datasetCruces.map((item) => (
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
    </section>
  );
}

export default PseConciliationSection;
