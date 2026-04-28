function Alerts({ alertas }) {
  if (!alertas || alertas.length === 0) {
    return null;
  }

  return (
    <div
      style={{
        display: "grid",
        gap: 8,
        marginBottom: 14,
      }}
    >
      {alertas.map((alerta, index) => (
        <div
          key={`${alerta}-${index}`}
          style={{
            background: "#fff4e6",
            border: "1px solid #ffd8a8",
            color: "#9c6644",
            borderRadius: 12,
            padding: "10px 12px",
            fontWeight: 600,
          }}
        >
          {alerta}
        </div>
      ))}
    </div>
  );
}

export default Alerts;
