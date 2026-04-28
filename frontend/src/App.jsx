import { useMemo, useState } from "react";
import Dashboard from "./components/Dashboard";

const API_BASE = (import.meta.env.VITE_API_URL || "http://localhost:8000").replace(/\/$/, "");

function decodeBase64ToBlob(base64, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new Blob([bytes], { type: mimeType });
}

function App() {
  const [selectedFile, setSelectedFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const filename = useMemo(() => {
    if (!selectedFile) return "Sin archivo seleccionado";
    return selectedFile.name;
  }, [selectedFile]);

  const handleFileChange = (event) => {
    const file = event.target.files?.[0] || null;
    setSelectedFile(file);
    setError("");
  };

  const triggerDownload = (base64Content, sourceName) => {
    const blob = decodeBase64ToBlob(
      base64Content,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const safeName = sourceName?.replace(/\.xlsx$/i, "") || "conciliado";
    link.href = url;
    link.download = `${safeName}_conciliado.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const handleProcess = async () => {
    if (!selectedFile) {
      setError("Selecciona un archivo .xlsx antes de procesar.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const response = await fetch(`${API_BASE}/procesar`, {
        method: "POST",
        body: formData,
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload?.detail || "No se pudo procesar el archivo");
      }

      setResult(payload);
      if (payload.file) {
        triggerDownload(payload.file, selectedFile.name);
      }
    } catch (err) {
      setError(err.message || "Error inesperado durante el procesamiento");
      setResult(null);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app-shell">
      <style>{`
        :root {
          --bg-a: #f2efe7;
          --bg-b: #dff0d8;
          --ink: #132a13;
          --ink-soft: #31572c;
          --card: #ffffffcc;
          --line: #8fb996;
          --accent: #4f772d;
          --accent-2: #90a955;
          --danger: #b02a37;
          --warn: #9a6700;
        }

        * { box-sizing: border-box; }

        body {
          margin: 0;
          font-family: "Sora", sans-serif;
          background:
            radial-gradient(circle at 10% 15%, #ffe8b6 0%, transparent 35%),
            radial-gradient(circle at 90% 75%, #c7f9cc 0%, transparent 40%),
            linear-gradient(145deg, var(--bg-a), var(--bg-b));
          color: var(--ink);
          min-height: 100vh;
        }

        .app-shell {
          min-height: 100vh;
          padding: 24px;
        }

        .container {
          max-width: 1200px;
          margin: 0 auto;
          display: grid;
          gap: 20px;
        }

        .hero {
          background: var(--card);
          border: 1px solid #ffffff;
          border-radius: 18px;
          padding: 22px;
          backdrop-filter: blur(8px);
          box-shadow: 0 15px 35px #1b433220;
        }

        .title {
          margin: 0;
          font-family: "Space Grotesk", sans-serif;
          font-size: clamp(1.4rem, 3vw, 2rem);
          letter-spacing: -0.02em;
        }

        .subtitle {
          margin-top: 8px;
          color: var(--ink-soft);
        }

        .uploader {
          margin-top: 18px;
          display: flex;
          flex-wrap: wrap;
          gap: 10px;
          align-items: center;
        }

        .input-file {
          border: 1px solid var(--line);
          border-radius: 10px;
          padding: 8px;
          background: #fff;
          max-width: 330px;
        }

        .filename {
          color: var(--ink-soft);
          font-size: 0.9rem;
          min-width: 220px;
        }

        .btn {
          border: 0;
          border-radius: 12px;
          padding: 10px 16px;
          font-family: "Space Grotesk", sans-serif;
          font-weight: 700;
          cursor: pointer;
          transition: transform 140ms ease, box-shadow 180ms ease;
        }

        .btn-primary {
          color: #fff;
          background: linear-gradient(120deg, var(--accent), var(--accent-2));
          box-shadow: 0 8px 18px #4f772d44;
        }

        .btn:disabled {
          cursor: not-allowed;
          filter: grayscale(20%);
          opacity: 0.75;
        }

        .btn:not(:disabled):hover {
          transform: translateY(-1px);
        }

        .error {
          margin-top: 12px;
          color: var(--danger);
          font-weight: 600;
        }

        @media (max-width: 768px) {
          .app-shell { padding: 14px; }
          .uploader { align-items: stretch; }
          .input-file { width: 100%; max-width: none; }
        }
      `}</style>

      <main className="container">
        <section className="hero">
          <h1 className="title">Conciliador Contable Inteligente</h1>
          <p className="subtitle">
            Carga tu archivo Excel, procesa en memoria y descarga el resultado conciliado sin persistencia.
          </p>

          <div className="uploader">
            <input
              className="input-file"
              type="file"
              accept=".xlsx"
              onChange={handleFileChange}
              disabled={loading}
            />
            <button className="btn btn-primary" onClick={handleProcess} disabled={loading}>
              {loading ? "Procesando..." : "Procesar archivo"}
            </button>
            <span className="filename">{filename}</span>
          </div>

          {error && <p className="error">{error}</p>}
        </section>

        <Dashboard data={result} loading={loading} />
      </main>
    </div>
  );
}

export default App;
