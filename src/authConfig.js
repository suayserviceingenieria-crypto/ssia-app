// ============================================================================
// CONFIGURACIÓN DE MICROSOFT ENTRA ID (MSAL) — S&S IA
// ============================================================================
// Esta configuración conecta la app con Microsoft Graph para leer y escribir
// en los libros de Excel de OneDrive. Es una capa DISTINTA e independiente
// del login por usuario/clave (SHA-256) que ya usan los usuarios operativos
// para entrar a la app día a día — ver src/utils/auth.js para eso. Aquí solo
// se trata de la conexión de la app con Microsoft 365.
// ============================================================================

import { LogLevel } from "@azure/msal-browser";

// Datos de tu registro en Entra ID (App registrations). El clientId y el
// tenantId de una SPA NO son secretos — están pensados para ser públicos,
// por eso es seguro que vivan aquí en el código del frontend. La seguridad
// real de este flujo (Authorization Code + PKCE) depende de que la URI de
// redirección esté registrada exactamente como "SPA" en Azure, tal como ya
// la configuraste.
const CLIENT_ID = "9af48a08-a47e-4671-9165-647ab49fa04d";
const TENANT_ID = "c3acf22c-148c-4c94-8c59-4a577a4af969";

export const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    // Debe coincidir EXACTO (protocolo, dominio, puerto Y RUTA) con la URI de
    // redirección tipo SPA que registraste en Azure. window.location.origin
    // por sí solo NO incluye la subcarpeta "/ssia-app/" donde vive la app en
    // GitHub Pages — por eso se combina con import.meta.env.BASE_URL, que
    // toma automáticamente el mismo "base" que configuramos en
    // vite.config.js, tanto en local (npm run dev) como en producción.
    redirectUri: window.location.origin + import.meta.env.BASE_URL,
    postLogoutRedirectUri: window.location.origin + import.meta.env.BASE_URL,
  },
  cache: {
    // localStorage, no sessionStorage: así la conexión con Microsoft
    // persiste entre sesiones (cerrar y volver a abrir el navegador) — el
    // login diario por usuario/clave sigue cerrándose solo (sessionStorage,
    // en utils/auth.js), pero la conexión con OneDrive necesita sobrevivir
    // más tiempo para poder reconectarse en silencio sin pedirle a cada
    // persona que la reconecte cada vez que abre la app.
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        switch (level) {
          case LogLevel.Error:
            console.error("[MSAL]", message);
            return;
          case LogLevel.Warning:
            console.warn("[MSAL]", message);
            return;
          default:
            return; // Info/Verbose silenciados — actívalos manualmente si necesitas depurar
        }
      },
    },
  },
};

// Permisos de Microsoft Graph que la app solicita. Deben coincidir con los
// que habilitaste en Azure (Files.ReadWrite, Files.ReadWrite.All y ahora
// también Calendars.ReadWrite, agregado para crear tareas de seguimiento
// comercial como eventos reales en el Calendario de Microsoft 365 de
// comercial@suaservice.com). Al agregar un scope nuevo, Azure Portal → tu
// App registration → API permissions → Add a permission → Microsoft Graph
// → Delegated permissions → Calendars.ReadWrite (y "Grant admin consent" si
// tu organización lo exige) — y cada persona va a tener que volver a
// conectar Microsoft una vez (reconectar/aceptar el nuevo permiso), no solo
// seguir usando la sesión que ya tenía.
export const graphScopes = ["Files.ReadWrite", "Files.ReadWrite.All", "Calendars.ReadWrite"];

export const graphBaseUrl = "https://graph.microsoft.com/v1.0";
