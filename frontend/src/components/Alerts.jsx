function Alerts({ alertas }) {
  if (!alertas || alertas.length === 0) {
    return null;
  }

  return (
    <div className="alerts-stack">
      {alertas.map((alerta, index) => (
        <div
          key={`${alerta}-${index}`}
          className="alert-card"
        >
          <strong>Alerta</strong>
          {alerta}
        </div>
      ))}
    </div>
  );
}

export default Alerts;
