import Alerts from "./Alerts";
import LogsTable from "./LogsTable";
import SummaryCards from "./SummaryCards";

function Dashboard({ data, loading }) {
  const resumen = data?.resumen || { cruzados: 0, posibles: 0, precision_estimada: 0 };
  const logs = data?.logs || [];
  const alertas = data?.alertas || [];

  return (
    <section>
      <SummaryCards resumen={resumen} loading={loading} />
      <Alerts alertas={alertas} />
      <LogsTable logs={logs} />
    </section>
  );
}

export default Dashboard;
