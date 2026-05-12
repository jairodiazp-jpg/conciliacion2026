function Card({ title, value, tone, subtitle }) {
  return (
    <article
      className="summary-card"
      style={{
        "--tone": tone,
      }}
    >
      <span className="summary-accent" />
      <p className="summary-label">{title}</p>
      <h2 className="summary-value">{value}</h2>
      {subtitle && <p className="summary-subtitle">{subtitle}</p>}
    </article>
  );
}

function SummaryCards({ resumen, loading }) {
  const precision = `${Math.round((Number(resumen.precision_estimada || 0) * 100))}%`;
  const porcentajeCruzados = loading ? "..." : `${resumen.porcentaje_cruzados || 0}% del total`;
  const porcentajePosibles = loading ? "..." : `${resumen.porcentaje_posibles || 0}% del total`;

  return (
    <div className="summary-grid">
      <Card 
        title="Total cruzados" 
        value={loading ? "..." : resumen.cruzados} 
        tone="#2563eb"
        subtitle={porcentajeCruzados}
      />
      <Card 
        title="Total posibles" 
        value={loading ? "..." : resumen.posibles} 
        tone="#f59e0b"
        subtitle={porcentajePosibles}
      />
      <Card 
        title="Precisión estimada" 
        value={loading ? "..." : precision} 
        tone="#21c4c4"
      />
    </div>
  );
}

export default SummaryCards;
