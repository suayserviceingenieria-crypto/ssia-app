import { useState } from "react";
import { verificarClave, guardarSesion, sincronizarSesionSupabase, migrarHashSiHaceFalta } from "../utils/auth";

/**
 * Pantalla de login. Bloquea el resto de la interfaz hasta que haya una
 * sesión válida — el componente que la usa (ej. App.jsx) decide cuándo
 * mostrarla, típicamente cuando leerSesion() devuelve null.
 *
 * @param {{ usuarios: Array, onIngresar: (sesion: object) => void }} props
 *   usuarios: la lista actual de usuarios autorizados (incluye los creados
 *             desde "Usuarios y accesos", no solo USUARIOS_SEMILLA).
 *   onIngresar: se llama con el objeto de sesión ya guardado en
 *               sessionStorage, para que el componente padre actualice su
 *               propio estado y deje de mostrar el login.
 */
export default function Login({ usuarios, onIngresar }) {
  const [usuario, setUsuario] = useState("");
  const [clave, setClave] = useState("");
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState("");

  async function intentarIngresar() {
    if (!usuario.trim() || !clave) {
      setError("Ingresa usuario y clave.");
      return;
    }
    setCargando(true);
    setError("");

    const usuarioNormalizado = usuario.trim().toLowerCase();
    const encontrado = usuarios.find((u) => u.usuario === usuarioNormalizado);

    if (!encontrado) {
      setError("Usuario o clave incorrectos.");
      setCargando(false);
      return;
    }

    const claveValida = await verificarClave(clave, encontrado.passwordHash);
    setCargando(false);

    if (!claveValida) {
      setError("Usuario o clave incorrectos.");
      return;
    }

    // Abre (o crea, la primera vez) la sesión real de Supabase que exigen
    // las políticas de seguridad de las tablas — ver la nota en utils/auth.js.
    // Nunca bloquea el login si falla: la persona entra a la app igual.
    await sincronizarSesionSupabase(usuarioNormalizado, clave);

    // Si esta cuenta todavía tenía el hash del esquema anterior (sal fija
    // compartida), lo reemplaza en silencio por uno con sal propia. No se
    // espera (no bloquea el login) — la función ya maneja sus propios
    // errores sin interrumpir nada.
    migrarHashSiHaceFalta(encontrado.id, encontrado.passwordHash, clave);

    const sesion = guardarSesion(encontrado);
    onIngresar(sesion);
  }

  function manejarEnter(e) {
    if (e.key === "Enter") intentarIngresar();
  }

  return (
    <div className="flex min-h-screen flex-col items-center justify-center bg-white px-4">
      <div className="fixed left-0 top-0 h-[5px] w-full shrink-0 bg-amber-400" />
      <div className="w-full max-w-xs">
        <p className="mb-1 text-center text-2xl font-bold leading-none tracking-tight">
          <span className="text-green-900">S</span>
          <span className="text-amber-400">&amp;</span>
          <span className="text-lime-700">S</span>
          <span className="ml-1 text-slate-700">IA</span>
        </p>
        <p className="mb-6 text-center text-xs text-slate-400">S&amp;S Intelligent Administration</p>

        <div className="space-y-3 rounded-xl border border-slate-200 p-5 shadow-sm">
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-600">Usuario</label>
            <input
              value={usuario}
              onChange={(e) => setUsuario(e.target.value)}
              onKeyDown={manejarEnter}
              className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm"
              autoFocus
              autoComplete="username"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-600">Clave</label>
            <input
              type="password"
              value={clave}
              onChange={(e) => setClave(e.target.value)}
              onKeyDown={manejarEnter}
              className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm"
              autoComplete="current-password"
            />
          </div>
          {error && <p className="text-xs text-rose-600">{error}</p>}
          <button
            onClick={intentarIngresar}
            disabled={cargando}
            className="w-full rounded-lg bg-green-900 px-4 py-2.5 text-sm font-medium text-white hover:bg-green-950 disabled:opacity-60"
          >
            {cargando ? "Verificando…" : "Ingresar"}
          </button>
        </div>
      </div>
    </div>
  );
}
