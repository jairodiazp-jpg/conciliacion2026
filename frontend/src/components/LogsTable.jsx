function badgeStyle(tipo) {
  const base = {
    display: "inline-block",
    padding: "4px 8px",
    borderRadius: 999,
    fontSize: 12,
    fontWeight: 700,
    textTransform: "uppercase",
    letterSpacing: "0.03em",
  };

  if (tipo === "aprobacion") {
    return { ...base, background: "#dbeafe", color: "#1d4ed8" };
  }
  if (tipo === "valor_fecha") {
    return { ...base, background: "#dcfce7", color: "#047857" };
  }
  return { ...base, background: "#fef3c7", color: "#b45309" };
}

function formatDate(value) {
  if (!value) return "-";

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;

  return new Intl.DateTimeFormat("es-CO", {
    year: "numeric",
    month: "short",
    day: "2-digit",
  }).format(date);
}

function formatValue(value) {
  const number = Number(value);
  if (Number.isNaN(number)) return "-";
  return number.toLocaleString("es-CO", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function LogsTable({ logs }) {
  return (
    <div className="logs-panel">
      <header className="logs-header">
        <div>
          <p className="section-eyebrow">Detalle operativo</p>
          <h2>Logs de conciliación</h2>
        </div>
        <span className="logs-count">{logs.length} registros</span>
      </header>

      {logs.length === 0 ? (
        <p className="logs-empty">
          Aún no hay registros. Procesa un archivo para ver los cruces.
        </p>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table className="logs-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Tipo</th>
                <th>Fecha</th>
                <th>Valor</th>
                <th>Confianza</th>
                <th>Detalle completo</th>
              </tr>
            </thead>
            <tbody>
              {logs.map((log, index) => (
                <tr key={`${log.tipo}-${index}`}>
                  <td>{index + 1}</td>
                  <td>
                    <span style={badgeStyle(log.tipo)}>{log.tipo}</span>
                  </td>
                  <td>{formatDate(log.fecha)}</td>
                  <td className="logs-value">{formatValue(log.valor)}</td>
                  <td>{Math.round(Number(log.confianza || 0) * 100)}%</td>
                  <td className="logs-detail">{log.detalle}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default LogsTable;
