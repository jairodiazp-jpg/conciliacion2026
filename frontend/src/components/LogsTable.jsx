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
    return { ...base, background: "#ffe8cc", color: "#a66300" };
  }
  if (tipo === "valor_fecha") {
    return { ...base, background: "#d8f3dc", color: "#2d6a4f" };
  }
  return { ...base, background: "#fff3bf", color: "#7f6000" };
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
    <div
      style={{
        background: "#ffffffcc",
        border: "1px solid #ffffff",
        borderRadius: 16,
        overflow: "hidden",
        boxShadow: "0 10px 24px #1b433214",
      }}
    >
      <header
        style={{
          padding: "14px 16px",
          borderBottom: "1px solid #d8f3dc",
          fontFamily: "Space Grotesk, sans-serif",
          fontWeight: 700,
        }}
      >
        Logs de conciliacion
      </header>

      {logs.length === 0 ? (
        <p style={{ margin: 0, padding: 16, color: "#31572c" }}>
          Aun no hay registros. Procesa un archivo para ver los cruces.
        </p>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 980 }}>
            <thead>
              <tr style={{ textAlign: "left", background: "#f1faee", position: "sticky", top: 0 }}>
                <th style={{ padding: 12, whiteSpace: "nowrap" }}>#</th>
                <th style={{ padding: 12, whiteSpace: "nowrap" }}>Tipo</th>
                <th style={{ padding: 12, whiteSpace: "nowrap" }}>Fecha</th>
                <th style={{ padding: 12, whiteSpace: "nowrap" }}>Valor</th>
                <th style={{ padding: 12, whiteSpace: "nowrap" }}>Confianza</th>
                <th style={{ padding: 12 }}>Detalle completo</th>
              </tr>
            </thead>
            <tbody>
              {logs.map((log, index) => (
                <tr
                  key={`${log.tipo}-${index}`}
                  style={{ borderTop: "1px solid #e9f5db", background: index % 2 === 0 ? "#ffffff" : "#f8fbf5" }}
                >
                  <td style={{ padding: 12, color: "#5c7c4c", fontWeight: 700 }}>{index + 1}</td>
                  <td style={{ padding: 12 }}>
                    <span style={badgeStyle(log.tipo)}>{log.tipo}</span>
                  </td>
                  <td style={{ padding: 12, whiteSpace: "nowrap" }}>{formatDate(log.fecha)}</td>
                  <td style={{ padding: 12, whiteSpace: "nowrap", fontWeight: 600 }}>{formatValue(log.valor)}</td>
                  <td style={{ padding: 12, whiteSpace: "nowrap" }}>{Math.round(Number(log.confianza || 0) * 100)}%</td>
                  <td style={{ padding: 12, color: "#31572c", minWidth: 420, lineHeight: 1.5 }}>{log.detalle}</td>
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
