import Alerts from "./Alerts";
import LogsTable from "./LogsTable";
import SummaryCards from "./SummaryCards";

function Dashboard({ data, loading }) {
  const resumen = data?.resumen || { cruzados: 0, posibles: 0, precision_estimada: 0 };
  const logs = data?.logs || [];
  const alertas = data?.alertas || [];

  return (
    <section className="dashboard-grid">
      <SummaryCards resumen={resumen} loading={loading} />
      <Alerts alertas={alertas} />
      <LogsTable logs={logs} />
      {!loading && logs.length === 0 && (
        <div className="empty-state">
          <h3>No hay conciliaciones aún</h3>
          <p>
            Sube un Excel para ver aquí el resumen, los logs detallados y las alertas del proceso.
          </p>
        </div>
      )}
    </section>
  );
}

export default Dashboard;
