function Card({ title, value, tone }) {
  return (
    <article
      style={{
        background: "#ffffffcc",
        border: "1px solid #ffffff",
        borderLeft: `6px solid ${tone}`,
        borderRadius: 14,
        padding: 16,
        boxShadow: "0 8px 20px #1b433218",
      }}
    >
      <p style={{ margin: 0, color: "#31572c", fontSize: 13 }}>{title}</p>
      <h2
        style={{
          margin: "8px 0 0",
          fontFamily: "Space Grotesk, sans-serif",
          fontSize: 28,
          color: "#132a13",
        }}
      >
        {value}
      </h2>
    </article>
  );
}

function SummaryCards({ resumen, loading }) {
  const precision = `${Math.round((Number(resumen.precision_estimada || 0) * 100))}%`;

  return (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
        gap: 12,
        marginBottom: 14,
      }}
    >
      <Card title="Total cruzados" value={loading ? "..." : resumen.cruzados} tone="#4f772d" />
      <Card title="Total posibles" value={loading ? "..." : resumen.posibles} tone="#bc6c25" />
      <Card title="Precision estimada" value={loading ? "..." : precision} tone="#386641" />
    </div>
  );
}

export default SummaryCards;
