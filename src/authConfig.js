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
    // Debe coincidir EXACTO (mismo protocolo, dominio y puerto) con la URI
    // de redirección tipo SPA que registraste en Azure. Si despliegas en
    // varios dominios (ej. localhost para pruebas y tu dominio real en
    // producción), registra AMBAS URIs en Azure — window.location.origin
    // toma automáticamente la que corresponda en cada entorno.
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    // sessionStorage, igual criterio que el login por clave: la conexión
    // con Microsoft se cierra sola al cerrar la pestaña o el navegador.
    cacheLocation: "sessionStorage",
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
// que habilitaste en Azure (Files.ReadWrite y Files.ReadWrite.All).
export const graphScopes = ["Files.ReadWrite", "Files.ReadWrite.All"];

export const graphBaseUrl = "https://graph.microsoft.com/v1.0";
