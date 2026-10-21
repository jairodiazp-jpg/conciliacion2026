import { useMemo, useState } from "react";

const TYPE_META = {
  aprobacion: { label: "Aprobación", className: "type-aprobacion" },
  valor_fecha: { label: "Valor + Fecha", className: "type-valor-fecha" },
  adquirencia_cruzada: { label: "Adquirencia", className: "type-adquirencia" },
  posible: { label: "Posible", className: "type-posible" },
};

function getTypeMeta(tipo) {
  return TYPE_META[tipo] || { label: tipo || "Sin tipo", className: "type-default" };
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

function confidencePercent(confianza) {
  const numeric = Number(confianza || 0);
  if (Number.isNaN(numeric)) return 0;
  return Math.max(0, Math.min(100, Math.round(numeric * 100)));
}

function confidenceTone(percent) {
  if (percent >= 90) return "confidence-high";
  if (percent >= 70) return "confidence-mid";
  return "confidence-low";
}

function LogsTable({ logs }) {
  const [searchText, setSearchText] = useState("");
  const [selectedType, setSelectedType] = useState("all");
  const [minimumConfidence, setMinimumConfidence] = useState(0);
  const [sortBy, setSortBy] = useState("confidence_desc");
  const [expandedRows, setExpandedRows] = useState({});

  const typeOptions = useMemo(() => {
    const uniqueTypes = Array.from(new Set(logs.map((log) => log?.tipo).filter(Boolean)));
    return uniqueTypes;
  }, [logs]);

  const filteredLogs = useMemo(() => {
    const normalizedSearch = searchText.trim().toLowerCase();

    const base = logs.filter((log) => {
      const percent = confidencePercent(log?.confianza);
      const byConfidence = percent >= minimumConfidence;
      const byType = selectedType === "all" || log?.tipo === selectedType;
      const byText =
        normalizedSearch.length === 0 ||
        String(log?.detalle || "").toLowerCase().includes(normalizedSearch) ||
        String(log?.tipo || "").toLowerCase().includes(normalizedSearch);

      return byConfidence && byType && byText;
    });

    const sorted = [...base];
    if (sortBy === "confidence_desc") {
      sorted.sort((a, b) => confidencePercent(b?.confianza) - confidencePercent(a?.confianza));
    } else if (sortBy === "confidence_asc") {
      sorted.sort((a, b) => confidencePercent(a?.confianza) - confidencePercent(b?.confianza));
    } else if (sortBy === "date_desc") {
      sorted.sort((a, b) => new Date(b?.fecha || 0).getTime() - new Date(a?.fecha || 0).getTime());
    } else if (sortBy === "date_asc") {
      sorted.sort((a, b) => new Date(a?.fecha || 0).getTime() - new Date(b?.fecha || 0).getTime());
    }

    return sorted;
  }, [logs, searchText, selectedType, minimumConfidence, sortBy]);

  const dashboardMetrics = useMemo(() => {
    const total = logs.length;
    const visibles = filteredLogs.length;
    const highConfidence = filteredLogs.filter((log) => confidencePercent(log?.confianza) >= 90).length;
    const avgConfidence =
      visibles === 0
        ? 0
        : Math.round(
            filteredLogs.reduce((acc, log) => acc + confidencePercent(log?.confianza), 0) /
              visibles
          );

    return { total, visibles, highConfidence, avgConfidence };
  }, [logs, filteredLogs]);

  const toggleExpanded = (rowKey) => {
    setExpandedRows((previous) => ({
      ...previous,
      [rowKey]: !previous[rowKey],
    }));
  };

  return (
    <div className="logs-panel">
      <header className="logs-header">
        <div>
          <p className="section-eyebrow">Detalle operativo</p>
          <h2>Logs de conciliación</h2>
        </div>
        <span className="logs-count">{dashboardMetrics.visibles} de {dashboardMetrics.total}</span>
      </header>

      <section className="logs-toolbar" aria-label="Filtros de logs">
        <div className="logs-kpis">
          <article className="kpi-card">
            <p>Visibles</p>
            <strong>{dashboardMetrics.visibles}</strong>
          </article>
          <article className="kpi-card">
            <p>Confianza alta</p>
            <strong>{dashboardMetrics.highConfidence}</strong>
          </article>
          <article className="kpi-card">
            <p>Promedio confianza</p>
            <strong>{dashboardMetrics.avgConfidence}%</strong>
          </article>
        </div>

        <div className="logs-filters">
          <label className="filter-field" htmlFor="logs-search">
            Buscar
            <input
              id="logs-search"
              type="search"
              placeholder="Tipo o detalle"
              value={searchText}
              onChange={(event) => setSearchText(event.target.value)}
              className="filter-input"
            />
          </label>

          <label className="filter-field" htmlFor="logs-type">
            Tipo
            <select
              id="logs-type"
              value={selectedType}
              onChange={(event) => setSelectedType(event.target.value)}
              className="filter-select"
            >
              <option value="all">Todos</option>
              {typeOptions.map((typeValue) => (
                <option key={typeValue} value={typeValue}>
                  {getTypeMeta(typeValue).label}
                </option>
              ))}
            </select>
          </label>

          <label className="filter-field" htmlFor="logs-sort">
            Orden
            <select
              id="logs-sort"
              value={sortBy}
              onChange={(event) => setSortBy(event.target.value)}
              className="filter-select"
            >
              <option value="confidence_desc">Confianza alta a baja</option>
              <option value="confidence_asc">Confianza baja a alta</option>
              <option value="date_desc">Fecha más reciente</option>
              <option value="date_asc">Fecha más antigua</option>
            </select>
          </label>

          <label className="filter-field filter-range" htmlFor="logs-confidence">
            Confianza mínima: {minimumConfidence}%
            <input
              id="logs-confidence"
              type="range"
              min="0"
              max="100"
              step="5"
              value={minimumConfidence}
              onChange={(event) => setMinimumConfidence(Number(event.target.value))}
              className="filter-slider"
            />
          </label>
        </div>
      </section>

      {logs.length === 0 ? (
        <p className="logs-empty">
          Aún no hay registros. Procesa un archivo para ver los cruces.
        </p>
      ) : filteredLogs.length === 0 ? (
        <p className="logs-empty">
          No hay resultados con los filtros actuales. Ajusta búsqueda, tipo o confianza.
        </p>
      ) : (
        <div className="logs-table-wrap">
          <table className="logs-table logs-table-interactive">
            <thead>
              <tr>
                <th>#</th>
                <th>Tipo</th>
                <th>Fecha</th>
                <th>Valor</th>
                <th>Confianza</th>
                <th>Detalle</th>
                <th>Acción</th>
              </tr>
            </thead>
            <tbody>
              {filteredLogs.map((log, index) => {
                const rowKey = `${log.tipo || "tipo"}-${index}-${log.fecha || "sin-fecha"}`;
                const confidence = confidencePercent(log.confianza);
                const typeMeta = getTypeMeta(log.tipo);
                const isExpanded = Boolean(expandedRows[rowKey]);

                return (
                  <tr key={rowKey}>
                    <td>
                      <span className="log-row-id">{index + 1}</span>
                    </td>
                    <td>
                      <span className={`log-type-chip ${typeMeta.className}`}>{typeMeta.label}</span>
                    </td>
                    <td>{formatDate(log.fecha)}</td>
                    <td className="logs-value">{formatValue(log.valor)}</td>
                    <td>
                      <span className={`confidence-pill ${confidenceTone(confidence)}`}>{confidence}%</span>
                    </td>
                    <td>
                      <div className="log-detail-wrap">
                        <p className={`log-detail-text ${isExpanded ? "expanded" : "collapsed"}`}>
                          {log.detalle || "-"}
                        </p>
                      </div>
                    </td>
                    <td>
                      <button
                        type="button"
                        className="row-action-btn"
                        onClick={() => toggleExpanded(rowKey)}
                      >
                        {isExpanded ? "Ver menos" : "Ver más"}
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default LogsTable;
