// ============================================================================
// CONEXIÓN CON MICROSOFT 365 — S&S IA
// ============================================================================
// Botón de conectar/desconectar la cuenta corporativa de Microsoft, para
// que la app pueda leer y escribir en los libros de Excel de OneDrive vía
// Graph API. Es una conexión de la APP con Microsoft — distinta del login
// diario por usuario/clave que ya usan los usuarios operativos.
// ============================================================================

import { useState } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { graphScopes } from "../authConfig";

export default function ConexionMicrosoft() {
  const { instance, accounts } = useMsal();
  const estaConectado = useIsAuthenticated();
  const [conectando, setConectando] = useState(false);
  const [error, setError] = useState("");

  async function conectar() {
    setConectando(true);
    setError("");
    try {
      await instance.loginPopup({ scopes: graphScopes });
    } catch (e) {
      setError(e.message || "No se pudo conectar con Microsoft 365.");
    } finally {
      setConectando(false);
    }
  }

  async function desconectar() {
    await instance.logoutPopup({ account: accounts[0] });
  }

  return (
    <div className="rounded-xl border border-slate-200 bg-white p-4">
      <p className="mb-1 text-sm font-semibold">Conexión con Microsoft 365</p>
      <p className="mb-3 text-[11px] text-slate-500">
        Necesaria para sincronizar cotizaciones, registros y nómina con los libros de Excel de OneDrive.
      </p>

      {estaConectado ? (
        <div className="flex items-center justify-between rounded-lg bg-green-50 px-3 py-2">
          <div>
            <p className="text-xs font-medium text-green-800">Conectado como {accounts[0]?.username}</p>
            <p className="text-[10px] text-green-600">Los permisos de OneDrive están activos.</p>
          </div>
          <button onClick={desconectar} className="rounded border border-green-300 px-2 py-1 text-[11px] font-medium text-green-700 hover:bg-green-100">
            Desconectar
          </button>
        </div>
      ) : (
        <div>
          <button onClick={conectar} disabled={conectando} className="rounded-lg bg-green-700 px-4 py-2 text-sm font-medium text-white hover:bg-green-900 disabled:opacity-60">
            {conectando ? "Conectando…" : "Conectar Microsoft 365"}
          </button>
          {error && <p className="mt-2 text-xs text-rose-600">{error}</p>}
        </div>
      )}
    </div>
  );
}
