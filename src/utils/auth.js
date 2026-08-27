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

// Sal fija. Cámbiala por una propia antes de desplegar a producción —
// mientras más única sea, menos útil es cualquier tabla precalculada
// genérica contra este archivo específico.
const SALT = "SSIA-2026-x7Qz";

const SESSION_KEY = "ssia_sesion";

/**
 * Calcula el hash SHA-256 de (sal + clave) usando la API nativa del
 * navegador (crypto.subtle) — no depende de ninguna librería externa.
 * @param {string} clave - la clave en texto plano a hashear
 * @returns {Promise<string>} el hash en hexadecimal (64 caracteres)
 */
export async function hashClave(clave) {
  const datos = new TextEncoder().encode(SALT + clave);
  const bufferHash = await crypto.subtle.digest("SHA-256", datos);
  const bytes = Array.from(new Uint8Array(bufferHash));
  return bytes.map((b) => b.toString(16).padStart(2, "0")).join("");
}

/**
 * Compara una clave en texto plano contra un hash ya almacenado.
 * @param {string} claveIngresada - lo que la persona escribió en el login
 * @param {string} hashAlmacenado - el passwordHash guardado para ese usuario
 * @returns {Promise<boolean>}
 */
export async function verificarClave(claveIngresada, hashAlmacenado) {
  const hashCalculado = await hashClave(claveIngresada);
  return hashCalculado === hashAlmacenado;
}

// ============================================================================
// ROLES
// ============================================================================
// modulos: null = acceso a todos los módulos. Una lista = solo esos.
// soloLectura: true = puede ver todo pero no crear/editar/eliminar nada.
export const ROLES = [
  { id: "ESTRATEGICA", label: "Gestión Estratégica", descripcion: "Acceso completo — Alta Dirección", modulos: null, soloLectura: false },
  { id: "COMERCIAL", label: "Gestión Comercial", descripcion: "Cotizaciones y terceros", modulos: ["cotizaciones", "terceros"], soloLectura: false },
  { id: "EVALUACION", label: "Seguimiento y Evaluación", descripcion: "Ve todo, no edita — revisión gerencial", modulos: null, soloLectura: true },
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
// CREDENCIALES INICIALES EN TEXTO PLANO (solo para que pruebes de una vez;
// bórralas de aquí o cámbialas apenas hagas login la primera vez):
//
//   Usuario: admin        / Clave: Admin2026*        (Gestión Estratégica)
//   Usuario: comercial1   / Clave: Comercial2026*     (Gestión Comercial)
//
// Sus hashes (ya calculados con la sal de arriba, SHA-256(sal + clave)):
//   admin      -> 0990b24ea4538900510ee09c04241ab5c6670e63123d09dc68b354796ee644f5
//   comercial1 -> 2123398c14da27efbf30b3631b56c5089084d13c88de82880042670ecd2fba8
//
// Si cambias la sal (SALT arriba), estos hashes ya NO sirven — tendrías que
// recalcularlos. Puedes recalcular cualquier hash abriendo la consola del
// navegador (F12) y ejecutando, por ejemplo:
//   await hashClave("TuClaveNueva*")
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
  {
    id: "u2",
    usuario: "comercial1",
    nombreCompleto: "Vendedor 1",
    rolId: "COMERCIAL",
    passwordHash: "2123398c14da27efbf30b3631b56c5089084d13c88de82880042670ecd2fba8",
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

/** Cierra la sesión (logout). */
export function cerrarSesion() {
  sessionStorage.removeItem(SESSION_KEY);
}

/** Devuelve el objeto de rol completo (con permisos) para un rolId dado. */
export function obtenerRol(rolId) {
  return ROLES.find((r) => r.id === rolId) ?? null;
}
