// ============================================================================
// CONEXIÓN CON MICROSOFT 365 — S&S IA
// ============================================================================
// Botón de conectar/desconectar la cuenta corporativa de Microsoft, para
// que la app pueda leer y escribir en los libros de Excel de OneDrive vía
// Graph API. Es una conexión de la APP con Microsoft — distinta del login
// diario por usuario/clave que ya usan los usuarios operativos.
//
// IMPORTANTE — mismo Drive para todos: quien se conecte aquí determina a
// cuál OneDrive se suben los archivos (PDFs, filas de Excel), sin importar
// con qué usuario de la app (admin/comercial/gerencia) haya iniciado
// sesión. Para que todo caiga en el mismo lugar, TODOS deben conectarse
// aquí con la MISMA cuenta de Microsoft (ej. subgerencia@suaservice.com) —
// si alguien se conecta con su propia cuenta personal, sus archivos se
// irán a SU PROPIO OneDrive, no al compartido.
//
// Usa REDIRECCIÓN DE PÁGINA COMPLETA (loginRedirect) para el login manual,
// no popup: más lento visualmente, pero mucho más confiable entre
// navegadores. Además, intenta reconectar SOLA en segundo plano (sin
// redirigir ni pedir nada) cada vez que la app carga, aprovechando que la
// sesión ahora se guarda en localStorage y sobrevive a cerrar el
// navegador — solo si esa reconexión silenciosa falla (por ejemplo, la
// sesión ya expiró del todo) se le pide a la persona reconectar a mano.
// ============================================================================

import { useEffect, useState } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { graphScopes, graphBaseUrl } from "../authConfig";

// Hook reutilizable: intenta la reconexión silenciosa al montar, y expone
// el estado para que cualquier pantalla (no solo Calendario) pueda mostrar
// una alerta si no hay conexión.
export function useConexionMicrosoft() {
  const { instance, accounts } = useMsal();
  const estaConectado = useIsAuthenticated();
  const [intentando, setIntentando] = useState(true);
  const [fallo, setFallo] = useState(false);

  useEffect(() => {
    if (estaConectado) {
      setIntentando(false);
      setFallo(false);
      return;
    }
    let activo = true;
    instance
      .ssoSilent({ scopes: graphScopes, account: accounts[0] })
      .then(() => {
        if (activo) { setIntentando(false); setFallo(false); }
      })
      .catch(() => {
        // Falla en silencio esperado: no hay sesión previa de Microsoft
        // en este navegador todavía, o ya expiró del todo. No se fuerza
        // ningún redirect automático (eso sí interrumpiría a la persona
        // sin que lo pidiera) — solo se marca la alerta.
        if (activo) { setIntentando(false); setFallo(true); }
      });
    return () => { activo = false; };
    // Solo se reintenta cuando cambia si HAY cuenta conectada o no —
    // no en cada render.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [estaConectado]);

  return { estaConectado, intentando, fallo, cuenta: accounts[0]?.username };
}

// Banner de alerta, para poner en cualquier pantalla (no solo Calendario)
// cuando la conexión con Microsoft no está activa.
export function AlertaConexionMicrosoft() {
  const { instance } = useMsal();
  const { estaConectado, intentando, fallo } = useConexionMicrosoft();

  if (estaConectado || intentando) return null;

  return (
    <div className="mb-4 flex items-center justify-between rounded-lg border border-amber-300 bg-amber-50 px-4 py-3">
      <div>
        <p className="text-sm font-semibold text-amber-800">⚠ Sin conexión con Microsoft 365</p>
        <p className="text-xs text-amber-700">
          {fallo ? "No se pudo reconectar en silencio — hay que conectar de nuevo." : "No conectado."}{" "}
          Mientras tanto, nada se está sincronizando a OneDrive (Excel ni PDF).
        </p>
      </div>
      <button
        onClick={() => instance.loginRedirect({ scopes: graphScopes })}
        className="rounded-lg bg-amber-600 px-3 py-2 text-sm font-medium text-white hover:bg-amber-700"
      >
        Conectar ahora
      </button>
    </div>
  );
}

export default function ConexionMicrosoft() {
  const { instance, accounts } = useMsal();
  const { estaConectado, intentando, fallo } = useConexionMicrosoft();
  const [driveId, setDriveId] = useState(null);
  const [buscandoDriveId, setBuscandoDriveId] = useState(false);
  const [errorDriveId, setErrorDriveId] = useState(null);

  async function verMiDriveId() {
    setBuscandoDriveId(true);
    setErrorDriveId(null);
    setDriveId(null);
    try {
      // Token directo por MSAL, sin pasar por obtenerTokenGraph — esa
      // función exige que DRIVE_ID ya esté configurado, que es
      // precisamente lo que este botón ayuda a conseguir por primera vez.
      const solicitud = { scopes: graphScopes, account: accounts[0] };
      let accessToken;
      try {
        accessToken = (await instance.acquireTokenSilent(solicitud)).accessToken;
      } catch {
        accessToken = (await instance.acquireTokenPopup(solicitud)).accessToken;
      }
      const respuesta = await fetch(`${graphBaseUrl}/me/drive`, { headers: { Authorization: `Bearer ${accessToken}` } });
      const datos = await respuesta.json();
      if (!respuesta.ok) throw new Error(datos?.error?.message || `Graph API respondió ${respuesta.status}`);
      setDriveId(datos.id);
    } catch (error) {
      setErrorDriveId(error?.message || String(error));
    } finally {
      setBuscandoDriveId(false);
    }
  }

  function conectar() {
    // loginRedirect navega fuera de la página — no hay nada que esperar
    // aquí, el resultado se procesa en main.jsx cuando la app vuelve a
    // cargar tras el login.
    instance.loginRedirect({ scopes: graphScopes });
  }

  function desconectar() {
    instance.logoutRedirect({ account: accounts[0] });
  }

  return (
    <div className="rounded-xl border border-slate-200 bg-white p-4">
      <p className="mb-1 text-sm font-semibold">Conexión con Microsoft 365</p>
      <p className="mb-3 text-[11px] text-slate-500">
        Necesaria para sincronizar cotizaciones, registros y nómina con los libros de Excel de OneDrive.
        Conéctate con tu propio correo corporativo — no hace falta compartir ninguna clave; los archivos
        siempre se guardan en la carpeta central, siempre que la tengas compartida contigo con permiso de edición.
      </p>

      {estaConectado ? (
        <div>
          <div className="flex items-center justify-between rounded-lg bg-green-50 px-3 py-2">
            <div>
              <p className="text-xs font-medium text-green-800">Conectado como {accounts[0]?.username}</p>
              <p className="text-[10px] text-green-600">Los permisos de OneDrive están activos. La conexión intentará renovarse sola la próxima vez que abras la app.</p>
            </div>
            <button onClick={desconectar} className="rounded border border-green-300 px-2 py-1 text-[11px] font-medium text-green-700 hover:bg-green-100">
              Desconectar
            </button>
          </div>

          {/* Herramienta temporal de diagnóstico — para obtener el ID real
              del Drive de esta cuenta sin depender de Graph Explorer, que
              ha dado problemas. Se puede quitar una vez que DRIVE_ID en
              graphExcelService.js ya esté configurado y funcionando bien. */}
          <div className="mt-2 rounded-lg border border-dashed border-slate-300 p-2.5">
            <button onClick={verMiDriveId} disabled={buscandoDriveId} className="text-[11px] font-medium text-slate-500 underline hover:text-slate-700 disabled:opacity-50">
              {buscandoDriveId ? "Consultando…" : "Ver el ID de mi Drive (para configurar DRIVE_ID)"}
            </button>
            {driveId && (
              <div className="mt-1.5">
                <p className="text-[10px] text-slate-500">Copia este texto completo y pégalo en DRIVE_ID:</p>
                <input readOnly value={driveId} onClick={(e) => e.target.select()} className="mt-1 w-full rounded border border-slate-300 bg-slate-50 px-2 py-1 text-[11px] font-mono" />
              </div>
            )}
            {errorDriveId && <p className="mt-1.5 text-[11px] text-rose-600">No se pudo obtener: {errorDriveId}</p>}
          </div>
        </div>
      ) : intentando ? (
        <p className="text-xs text-slate-500">Intentando reconectar en silencio…</p>
      ) : (
        <div>
          {fallo && (
            <p className="mb-2 text-xs text-amber-600">
              No se pudo reconectar en silencio — hace falta conectar de nuevo a mano.
            </p>
          )}
          <button onClick={conectar} className="rounded-lg bg-green-700 px-4 py-2 text-sm font-medium text-white hover:bg-green-900">
            Conectar Microsoft 365
          </button>
        </div>
      )}
    </div>
  );
}
