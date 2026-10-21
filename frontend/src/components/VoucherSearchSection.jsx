import { useState } from "react";

function VoucherSearchSection({ apiBase }) {
  const [memorando, setMemorando] = useState(null);
  const [vouchers, setVouchers] = useState([]);
  const [value, setValue] = useState("");
  const [exactDate, setExactDate] = useState("");
  const [store, setStore] = useState("");
  const [reference, setReference] = useState("");
  const [account, setAccount] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const analyze = async () => {
    if (!memorando || vouchers.length === 0) {
      setError("Selecciona el memorando conciliado y al menos un soporte.");
      return;
    }
    setLoading(true);
    setError("");
    try {
      const form = new FormData();
      form.append("memorando_file", memorando);
      vouchers.forEach((file) => form.append("voucher_files", file));
      form.append("value", value);
      if (exactDate) form.append("exact_date", exactDate);
      if (store) form.append("store", store);
      if (reference) form.append("reference", reference);
      if (account) form.append("account", account);
      const response = await fetch(`${apiBase}/voucher/buscar`, { method: "POST", body: form });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(payload.detail || `No se pudo analizar (${response.status})`);
      setResult(payload);
    } catch (analysisError) {
      setError(analysisError.message || "No se pudieron analizar los documentos.");
      setResult(null);
    } finally {
      setLoading(false);
    }
  };

  const search = async () => {
    if (!value.trim()) {
      setError("El valor exacto es obligatorio.");
      return;
    }
    await analyze();
  };

  const clear = () => {
    setValue("");
    setExactDate("");
    setStore("");
    setReference("");
    setAccount("");
    setResult(null);
    setError("");
  };

  return (
    <section className="upload-panel">
      <div className="panel-heading">
        <span className="cloud-badge">⌕</span>
        <div>
          <h2>Buscar valor voucher</h2>
          <p>Consulta de solo lectura. Los documentos originales no se modifican.</p>
        </div>
      </div>
      <div className="drop-grid">
        <label className="drop-card has-file">
          <input type="file" accept=".xlsx,.xlsm" onChange={(event) => setMemorando(event.target.files?.[0] || null)} disabled={loading} />
          <strong>Memorando conciliado</strong>
          <span className="hint">Lectura exclusiva</span>
          <span className="filename">{memorando?.name || "Seleccionar Excel"}</span>
        </label>
        <label className="drop-card has-file">
          <input type="file" multiple accept=".xlsx,.xlsm,.csv,.pdf,.png,.jpg,.jpeg" onChange={(event) => setVouchers(Array.from(event.target.files || []))} disabled={loading} />
          <strong>Vouchers y soportes</strong>
          <span className="hint">Excel, CSV, PDF e imágenes</span>
          <span className="filename">{vouchers.length ? `${vouchers.length} archivo(s) seleccionado(s)` : "Seleccionar soportes"}</span>
        </label>
      </div>
      <div className="refs-grid">
        <div className="ref-field"><label>Valor exacto</label><input value={value} onChange={(event) => setValue(event.target.value)} placeholder="Ej. 28144.00" /></div>
        <div className="ref-field"><label>Fecha exacta (opcional)</label><input type="date" value={exactDate} onChange={(event) => setExactDate(event.target.value)} /></div>
        <div className="ref-field"><label>Punto de venta (opcional)</label><input value={store} onChange={(event) => setStore(event.target.value)} /></div>
        <div className="ref-field"><label>Referencia o aprobación</label><input value={reference} onChange={(event) => setReference(event.target.value)} /></div>
        <div className="ref-field"><label>Cuenta bancaria</label><input value={account} onChange={(event) => setAccount(event.target.value)} /></div>
      </div>
      <div className="action-row">
        <button className="btn btn-primary" onClick={analyze} disabled={loading || !memorando || vouchers.length === 0}>{loading ? "Analizando..." : "Analizar documentos"}</button>
        <button className="btn btn-secondary" onClick={search} disabled={loading || !value.trim() || !memorando || vouchers.length === 0}>Buscar valor</button>
        <button className="btn btn-ghost" onClick={clear} disabled={loading}>Limpiar búsqueda</button>
      </div>
      {error && <p className="error">{error}</p>}
      {result && (
        <div className="logs-panel">
          <header className="logs-header"><div><p className="section-eyebrow">Consulta de solo lectura</p><h2>Resultados</h2></div></header>
          <div className="summary-grid">
            {Object.entries(result.summary).map(([key, item]) => <div className="summary-card" key={key}><strong>{item}</strong><span>{key.replaceAll("_", " ")}</span></div>)}
          </div>
          {result.issues.map((issue, index) => <p className="error" key={`${issue.file}-${index}`}>{issue.file}: {issue.detail}</p>)}
          <div style={{ overflowX: "auto" }}>
            <table className="logs-table">
              <thead><tr><th>Estado</th><th>Valor encontrado</th><th>Fecha</th><th>Punto de venta</th><th>Archivo</th><th>Ubicación</th><th>Evidencia</th><th>Observación</th></tr></thead>
              <tbody>{result.results.map((item, index) => <tr key={`${item.file}-${item.row}-${index}`}><td>{item.status}</td><td>{item.value}</td><td>{item.date || "NO VERIFICADA"}</td><td>{item.store}</td><td>{item.file}</td><td>{item.location}</td><td className="logs-detail">{item.evidence}</td><td>{item.observation}</td></tr>)}</tbody>
            </table>
          </div>
        </div>
      )}
    </section>
  );
}

export default VoucherSearchSection;
