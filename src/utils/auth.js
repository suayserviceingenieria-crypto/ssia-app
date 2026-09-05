import { supabase } from "../supabase/supabaseClient";

// ============================================================================
// AUTENTICACIÓN — S&S IA
// ============================================================================
// Seguridad SIN backend: las claves nunca se guardan en texto plano en el
// código. Se guarda solo el hash SHA-256 (clave + sal) de cada una.
//
// ADVERTENCIA HONESTA SOBRE EL ALCANCE DE ESTA SEGURIDAD:
// Este hashing evita que alguien que abra el código fuente vea las claves
// reales a simple vista — es un control razonable para uso diario del
// equipo. NO es seguridad de nivel backend: como todo corre en el navegador,
// una persona con conocimientos técnicos podría inspeccionar el código,
// ver la lista de hashes, y con suficiente esfuerzo intentar un ataque de
// diccionario/fuerza bruta offline contra ellos (la sal fija ayuda contra
// tablas arcoíris genéricas, pero al ser fija y estar en el mismo código,
// no protege contra alguien que ataque específicamente este archivo).
// Cuando este proyecto conecte un backend real (ej. Supabase Auth), este
// módulo debe reemplazarse por autenticación server-side de verdad.
// ============================================================================

// Sal fija ANTERIOR — ya NO se usa para claves nuevas. Se conserva solo
// para poder seguir verificando (y migrar en silencio, ver
// migrarHashSiHaceFalta más abajo) las claves que ya estaban guardadas
// con el esquema viejo, de una sola sal compartida por todos.
const SALT_LEGADA = "SSIA-2026-x7Qz";

const SESSION_KEY = "ssia_sesion";

/** SHA-256 de un texto, en hexadecimal — usando crypto.subtle nativo. */
async function sha256Hex(texto) {
  const datos = new TextEncoder().encode(texto);
  const bufferHash = await crypto.subtle.digest("SHA-256", datos);
  const bytes = Array.from(new Uint8Array(bufferHash));
  return bytes.map((b) => b.toString(16).padStart(2, "0")).join("");
}

/** Hash con el esquema ANTERIOR (sal fija compartida) — solo para verificar claves ya guardadas así. */
async function hashClaveLegado(clave) {
  return sha256Hex(SALT_LEGADA + clave);
}

/** 16 bytes al azar, como texto hexadecimal (32 caracteres) — una sal distinta cada vez. */
function generarSal() {
  const bytes = crypto.getRandomValues(new Uint8Array(16));
  return Array.from(bytes).map((b) => b.toString(16).padStart(2, "0")).join("");
}

/**
 * Genera el hash de una clave con el esquema NUEVO: sal propia y
 * aleatoria en cada llamada, guardada junto con el hash en un solo texto
 * ("sal:hash"). Es lo que hay que usar para cualquier clave nueva o
 * cambiada — nunca produce el mismo resultado dos veces, ni entre
 * usuarios ni entre cambios de clave de la misma persona.
 * @param {string} clave - la clave en texto plano a hashear
 * @returns {Promise<string>} "sal:hash", listo para guardar en password_hash
 */
export async function generarHashClave(clave) {
  const sal = generarSal();
  const hash = await sha256Hex(sal + clave);
  return `${sal}:${hash}`;
}

/**
 * Compara una clave en texto plano contra un hash ya almacenado. Entiende
 * los dos formatos que pueden existir en la base de datos: el nuevo
 * ("sal:hash", una sal propia por clave) y el anterior (un solo hash con
 * la sal fija compartida) — así ninguna clave que ya existía antes de
 * este cambio deja de funcionar.
 * @param {string} claveIngresada - lo que la persona escribió en el login
 * @param {string} hashAlmacenado - el passwordHash guardado para ese usuario
 * @returns {Promise<boolean>}
 */
export async function verificarClave(claveIngresada, hashAlmacenado) {
  if (typeof hashAlmacenado === "string" && hashAlmacenado.includes(":")) {
    const [sal, hash] = hashAlmacenado.split(":");
    const hashCalculado = await sha256Hex(sal + claveIngresada);
    return hashCalculado === hash;
  }
  const hashCalculado = await hashClaveLegado(claveIngresada);
  return hashCalculado === hashAlmacenado;
}

/**
 * Si el hash guardado todavía usa el esquema anterior (sal fija
 * compartida), lo reemplaza en Supabase por uno nuevo con sal propia —
 * en silencio, la próxima vez que esa persona entra con su clave
 * correcta. Si algo falla (ej. sin conexión) simplemente no migra nada
 * todavía y se vuelve a intentar en el siguiente login — nunca interrumpe
 * el ingreso a la app.
 * @param {string} usuarioId - id del usuario en la tabla "usuarios"
 * @param {string} hashActual - el passwordHash que ya tenía guardado
 * @param {string} claveEnTextoPlano - la clave que la persona acaba de escribir (ya verificada)
 */
export async function migrarHashSiHaceFalta(usuarioId, hashActual, claveEnTextoPlano) {
  if (typeof hashActual === "string" && hashActual.includes(":")) return; // ya está en el esquema nuevo
  try {
    const nuevoHash = await generarHashClave(claveEnTextoPlano);
    const { error } = await supabase.from("usuarios").update({ password_hash: nuevoHash }).eq("id", usuarioId);
    if (error) console.error("[Seguridad] No se pudo migrar el hash de la clave a sal propia:", error);
  } catch (e) {
    console.error("[Seguridad] No se pudo migrar el hash de la clave a sal propia:", e);
  }
}

// ============================================================================
// ROLES
// ============================================================================
// modulos: null = acceso a todos los módulos. Una lista = solo esos.
// soloLectura: true = puede ver todo pero no crear/editar/eliminar nada,
// EXCEPTO en los módulos listados en modulosEditables (si los hay) — ahí
// sí puede crear/editar/eliminar como si no tuviera solo lectura.
export const ROLES = [
  { id: "ESTRATEGICA", label: "Gestión Estratégica", descripcion: "Acceso completo — Alta Dirección", modulos: null, soloLectura: false },
  { id: "COMERCIAL", label: "Gestión Comercial", descripcion: "Cotizaciones y terceros", modulos: ["comercial"], soloLectura: false },
  { id: "EVALUACION", label: "Seguimiento y Evaluación", descripcion: "Ve todo; puede elaborar cotizaciones y gestionar clientes — el resto es de solo consulta", modulos: null, soloLectura: true, modulosEditables: ["cotizaciones", "terceros"] },
];

// ============================================================================
// USUARIOS AUTORIZADOS
// ============================================================================
// Cada usuario guarda passwordHash (no la clave). Este arreglo es el punto
// de partida — en la app, Gestión Estratégica puede crear/editar/eliminar
// usuarios desde la pantalla "Usuarios y accesos"; esos cambios viven en
// memoria de React (useState), así que se pierden al recargar la página a
// menos que conectes persistencia real (ej. Supabase) más adelante.
//
// ----------------------------------------------------------------------
// CREDENCIAL INICIAL EN TEXTO PLANO (solo para que puedas entrar la primera
// vez; cámbiala apenas hagas login, desde "Usuarios y accesos"):
//
//   Usuario: admin  / Clave: Admin2026*  (Gestión Estratégica)
//
// Su hash (calculado con el esquema anterior, de sal fija — se migra solo
// a sal propia la primera vez que ese usuario inicia sesión):
//   admin -> 0990b24ea4538900510ee09c04241ab5c6670e63123d09dc68b354796ee644f5
//
// Con esta única cuenta, Gestión Estratégica crea desde la app al resto del
// equipo (comerciales, seguimiento y evaluación) — no hace falta dejar más
// usuarios de ejemplo aquí en el código.
//
// Para recalcular un hash a mano (esquema nuevo, con sal propia), abre la
// consola del navegador (F12) y ejecuta, por ejemplo:
//   await generarHashClave("TuClaveNueva*")
// (con este módulo importado en el contexto, o pegando la función completa).
// ----------------------------------------------------------------------
export const USUARIOS_SEMILLA = [
  {
    id: "u1",
    usuario: "admin",
    nombreCompleto: "Administrador",
    rolId: "ESTRATEGICA",
    passwordHash: "0990b24ea4538900510ee09c04241ab5c6670e63123d09dc68b354796ee644f5",
  },
];

// ============================================================================
// SESIÓN (sessionStorage — se borra sola al cerrar la pestaña o el navegador)
// ============================================================================

/**
 * Guarda la sesión activa. Incluye el rol para poder restringir funciones
 * sin tener que volver a consultar la lista de usuarios en cada pantalla.
 */
export function guardarSesion(usuario) {
  const rol = ROLES.find((r) => r.id === usuario.rolId);
  const sesion = {
    id: usuario.id,
    usuario: usuario.usuario,
    nombreCompleto: usuario.nombreCompleto,
    rolId: usuario.rolId,
    rolLabel: rol?.label ?? usuario.rolId,
  };
  sessionStorage.setItem(SESSION_KEY, JSON.stringify(sesion));
  return sesion;
}

/** Lee la sesión activa desde sessionStorage. null si no hay ninguna. */
export function leerSesion() {
  const crudo = sessionStorage.getItem(SESSION_KEY);
  if (!crudo) return null;
  try {
    return JSON.parse(crudo);
  } catch {
    return null;
  }
}

/** Cierra la sesión (logout) — también cierra la sesión real de Supabase. */
export function cerrarSesion() {
  sessionStorage.removeItem(SESSION_KEY);
  supabase.auth.signOut().catch(() => {}); // best-effort, no bloquea el logout
}

/** Devuelve el objeto de rol completo (con permisos) para un rolId dado. */
export function obtenerRol(rolId) {
  return ROLES.find((r) => r.id === rolId) ?? null;
}

// ============================================================================
// SESIÓN REAL EN SUPABASE — para que Row Level Security pueda exigir
// "authenticated" en vez de dejar las tablas abiertas a cualquiera
// ============================================================================
// El login de arriba (usuario/clave + SHA-256) sigue siendo el que decide
// si la persona entra a la app — esto NO lo reemplaza. Es un paso adicional:
// usa la misma clave que la persona ya escribió para abrir (o crear, la
// primera vez que entra después de este cambio) una sesión real de
// Supabase Auth, mapeando "usuario" a un correo interno ficticio
// ("usuario@ssia.local") que nunca se usa para enviar nada — Supabase Auth
// simplemente lo exige como formato.
//
// Si falla, la persona IGUAL entra a la app con normalidad (nunca bloquea
// el login) — pero sus lecturas/escrituras a Supabase fallarán una vez
// actives RLS "authenticated" en las tablas, hasta resolver el caso (ver
// la nota junto a actualizarUsuario en App.jsx para cuando un administrador
// le resetea la clave a otra persona).
export async function sincronizarSesionSupabase(usuario, clave) {
  const correo = `${usuario}@ssia.local`;

  const { error: errorIngreso } = await supabase.auth.signInWithPassword({ email: correo, password: clave });
  if (!errorIngreso) return { ok: true };

  // No existe todavía la cuenta de Supabase Auth para este usuario (primera
  // vez que entra tras este cambio) — la creamos ahora, con esta misma clave.
  const { error: errorRegistro } = await supabase.auth.signUp({ email: correo, password: clave });
  if (!errorRegistro) return { ok: true };

  console.error("[Supabase Auth] No se pudo abrir/crear la sesión real:", errorRegistro || errorIngreso);
  return { ok: false, error: errorRegistro || errorIngreso };
}
