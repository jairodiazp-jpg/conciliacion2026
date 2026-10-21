import { useState } from "react";

function downloadBase64(fileContent, name) {
  const binary = atob(fileContent);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  const url = URL.createObjectURL(new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  }));
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function FilePicker({ title, file, onChange, onRemove, disabled }) {
  return (
    <div className="drop-card has-file">
      <input
        type="file"
        accept=".xlsx"
        aria-label={title}
        disabled={disabled}
        onChange={(event) => onChange(event.target.files?.[0] || null)}
      />
      <strong>{title}</strong>
      <span className="hint">Archivo .xlsx requerido</span>
      <span className="filename">{file ? file.name : "Selecciona o arrastra un archivo"}</span>
      {file && (
        <button type="button" className="btn btn-ghost" onClick={onRemove} disabled={disabled}>
          Retirar archivo
        </button>
      )}
    </div>
  );
}

function QueryMemorandoSection({ apiBase }) {
  const [queryFile, setQueryFile] = useState(null);
  const [memorandoFile, setMemorandoFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const process = async () => {
    if (!queryFile || !memorandoFile) {
      setError("Selecciona el archivo Query y el memorando ya conciliado.");
      return;
    }
    setLoading(true);
    setError("");
    try {
      const formData = new FormData();
      formData.append("query_file", queryFile);
      formData.append("memorando_file", memorandoFile);
      const response = await fetch(`${apiBase}/query-memorando/conciliar`, { method: "POST", body: formData });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(payload.detail || `No se pudo procesar (${response.status})`);
      setResult(payload);
    } catch (processingError) {
      setError(processingError.message || "Error procesando el Query.");
      setResult(null);
    } finally {
      setLoading(false);
    }
  };

  const clear = () => {
    setQueryFile(null);
    setMemorandoFile(null);
    setResult(null);
    setError("");
  };

  return (
    <section className="upload-panel">
      <div className="panel-heading">
        <span className="cloud-badge">Q</span>
        <div>
          <h2>Conciliar Query y memorando ya conciliado</h2>
          <p>Proceso independiente con coincidencia exacta de cuenta, valor y fecha.</p>
        </div>
      </div>
      <div className="drop-grid">
        <FilePicker title="Archivo Query" file={queryFile} onChange={setQueryFile} onRemove={() => setQueryFile(null)} disabled={loading} />
        <FilePicker title="Memorando ya conciliado" file={memorandoFile} onChange={setMemorandoFile} onRemove={() => setMemorandoFile(null)} disabled={loading} />
      </div>
      <div className="action-row">
        <button className="btn btn-primary" onClick={process} disabled={loading}>
          {loading ? "Validando..." : "Ejecutar conciliación Query"}
        </button>
        <button className="btn btn-secondary" onClick={() => result && downloadBase64(result.query_file.file, result.query_file.name)} disabled={!result}>
          Descargar Query
        </button>
        <button className="btn btn-secondary" onClick={() => result && downloadBase64(result.memorando_file.file, result.memorando_file.name)} disabled={!result}>
          Descargar memorando
        </button>
        <button className="btn btn-ghost" onClick={clear} disabled={loading}>Limpiar</button>
      </div>
      {error && <p className="error">{error}</p>}
      {result && (
        <div className="logs-panel">
          <header className="logs-header">
            <div><p className="section-eyebrow">Resultado independiente</p><h2>Resumen de conciliación</h2></div>
          </header>
          <div className="summary-grid">
            {Object.entries(result.summary).map(([key, value]) => <div className="summary-card" key={key}><strong>{value}</strong><span>{key.replaceAll("_", " ")}</span></div>)}
          </div>
          {result.issues.length > 0 && <p className="error">Se detectaron {result.issues.length} problemas de validación. No se inventaron datos.</p>}
          <div style={{ overflowX: "auto" }}>
            <table className="logs-table">
              <thead><tr><th>Estado</th><th>Query</th><th>Memorando</th><th>Motivo</th></tr></thead>
              <tbody>
                {result.results.map((item, index) => (
                  <tr key={`${item.query?.sheet}-${item.query?.row}-${index}`}>
                    <td>{item.status}</td>
                    <td>{item.query ? `${item.query.sheet} fila ${item.query.row}` : "-"}</td>
                    <td>{item.memorando ? `${item.memorando.sheet} fila ${item.memorando.row}` : item.candidates?.map((candidate) => `${candidate.sheet} fila ${candidate.row}`).join(", ") || "-"}</td>
                    <td className="logs-detail">{item.reason}</td>
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

export default QueryMemorandoSection;
