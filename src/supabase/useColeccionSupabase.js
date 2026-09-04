// ============================================================================
// DATOS EN TIEMPO REAL — S&S IA
// ============================================================================
// useColeccionSupabase reemplaza a usePersistido (localStorage): en vez de
// guardar solo en el navegador de cada quien, lee y escribe directo en la
// base de datos compartida — y se suscribe a los cambios que haga
// CUALQUIER dispositivo conectado, para que todos vean lo mismo sin tener
// que refrescar la página.
//
// Uso típico (reemplaza una línea "usePersistido" por esta):
//   const { registros: clientes, crear: crearCliente,
//           actualizar: actualizarCliente, eliminar: eliminarCliente }
//     = useColeccionSupabase("clientes");
// ============================================================================

import { useState, useEffect, useCallback } from "react";
import { supabase } from "./supabaseClient";

/**
 * @param {string} tabla - nombre de la tabla en Supabase (ej. "clientes").
 * @param {object} opciones
 *   @param {(fila: object) => object} opciones.desdeFila - transforma una
 *     fila de la base de datos al formato que usa la app. Por defecto,
 *     si la fila tiene una columna "datos" (jsonb), la aplana junto al id.
 *   @param {(registro: object) => object} opciones.haciaFila - transforma
 *     un registro de la app al formato de fila para guardar. Por defecto,
 *     envuelve todo menos el id dentro de una columna "datos" (jsonb).
 */
export function useColeccionSupabase(tabla, opciones = {}) {
  const desdeFila = opciones.desdeFila || ((fila) => (fila.datos ? { id: fila.id, ...fila.datos } : fila));
  const haciaFila = opciones.haciaFila || ((registro) => {
    const { id, ...resto } = registro;
    return { id, datos: resto };
  });

  const [registros, setRegistros] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    let activo = true;

    async function cargarInicial() {
      const { data, error: errorCarga } = await supabase.from(tabla).select("*");
      if (!activo) return;
      if (errorCarga) {
        console.error(`[Supabase] Error cargando "${tabla}":`, errorCarga);
        setError(errorCarga);
      } else {
        setRegistros((data || []).map(desdeFila));
      }
      setCargando(false);
    }
    cargarInicial();

    // Suscripción en tiempo real: cualquier INSERT/UPDATE/DELETE que
    // ocurra en esta tabla (sin importar quién ni desde qué dispositivo lo
    // haga) llega aquí y actualiza el estado local automáticamente.
    const canal = supabase
      .channel(`cambios-${tabla}`)
      .on("postgres_changes", { event: "*", schema: "public", table: tabla }, (payload) => {
        setRegistros((actuales) => {
          if (payload.eventType === "INSERT") {
            const nuevo = desdeFila(payload.new);
            if (actuales.some((r) => r.id === nuevo.id)) return actuales; // ya lo teníamos (eco de nuestra propia escritura)
            return [...actuales, nuevo];
          }
          if (payload.eventType === "UPDATE") {
            const actualizado = desdeFila(payload.new);
            return actuales.map((r) => (r.id === actualizado.id ? actualizado : r));
          }
          if (payload.eventType === "DELETE") {
            return actuales.filter((r) => r.id !== payload.old.id);
          }
          return actuales;
        });
      })
      .subscribe();

    return () => {
      activo = false;
      supabase.removeChannel(canal);
    };
  }, [tabla]);

  const crear = useCallback(async (registro) => {
    const { error: errorEscritura } = await supabase.from(tabla).insert(haciaFila(registro));
    if (errorEscritura) console.error(`[Supabase] Error creando en "${tabla}":`, errorEscritura);
    return !errorEscritura;
  }, [tabla]);

  const actualizar = useCallback(async (id, cambios) => {
    const actual = registros.find((r) => r.id === id);
    if (!actual) return false;
    const { error: errorEscritura } = await supabase.from(tabla).update(haciaFila({ ...actual, ...cambios })).eq("id", id);
    if (errorEscritura) console.error(`[Supabase] Error actualizando en "${tabla}":`, errorEscritura);
    return !errorEscritura;
  }, [tabla, registros]);

  const eliminar = useCallback(async (id) => {
    const { error: errorEscritura } = await supabase.from(tabla).delete().eq("id", id);
    if (errorEscritura) console.error(`[Supabase] Error eliminando en "${tabla}":`, errorEscritura);
    return !errorEscritura;
  }, [tabla]);

  return { registros, setRegistros, crear, actualizar, eliminar, cargando, error };
}

/**
 * Para un valor único compartido (no una lista) — hoy solo el NIT de la
 * empresa, en la tabla "configuracion" (una sola fila, id=1). Funciona
 * igual que useState, pero leyendo/escribiendo esa fila en Supabase, con
 * el mismo tiempo real: si alguien lo cambia desde otro dispositivo, se
 * actualiza solo aquí también.
 */
export function useConfiguracion(campo) {
  const [valor, setValorLocal] = useState("");

  useEffect(() => {
    let activo = true;
    supabase.from("configuracion").select(campo).eq("id", 1).single().then(({ data }) => {
      if (activo && data) setValorLocal(data[campo] || "");
    });

    const canal = supabase
      .channel(`cambios-configuracion-${campo}`)
      .on("postgres_changes", { event: "UPDATE", schema: "public", table: "configuracion" }, (payload) => {
        setValorLocal(payload.new[campo] || "");
      })
      .subscribe();

    return () => {
      activo = false;
      supabase.removeChannel(canal);
    };
  }, [campo]);

  async function setValor(nuevo) {
    setValorLocal(nuevo); // optimista, para que la persona vea el cambio de inmediato
    const { error } = await supabase.from("configuracion").update({ [campo]: nuevo }).eq("id", 1);
    if (error) console.error(`[Supabase] Error guardando "${campo}" en configuracion:`, error);
  }

  return [valor, setValor];
}
