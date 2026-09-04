// ============================================================================
// PUNTO DE ENTRADA — S&S IA
// ============================================================================
// Envuelve toda la app en MsalProvider para que cualquier componente pueda
// usar los hooks de @azure/msal-react (useMsal, useIsAuthenticated, etc.)
// para la conexión con Microsoft 365 / OneDrive. Esto es independiente del
// login por usuario/clave (SHA-256) que ya maneja App.jsx internamente.
//
// Usa REDIRECCIÓN DE PÁGINA COMPLETA (loginRedirect), no popup — es el
// patrón más confiable entre navegadores, porque no depende de que el
// navegador mantenga la comunicación entre una ventana emergente y la
// principal (eso fue justo lo que causaba el error "timed_out" con popup).
// handleRedirectPromise() procesa la respuesta de Microsoft cuando el
// navegador vuelve a cargar la app después de que la persona inicia sesión.
// ============================================================================

import React from "react";
import ReactDOM from "react-dom/client";
import { PublicClientApplication, EventType } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";
import App from "./App";
import "./index.css";

const msalInstance = new PublicClientApplication(msalConfig);

async function iniciar() {
  // Requerido por msal-browser (v3 en adelante) antes de usar la instancia.
  await msalInstance.initialize();

  // Procesa la respuesta si la app se acaba de cargar de vuelta después de
  // un loginRedirect/logoutRedirect. Si no viene de una redirección de
  // Microsoft, esto simplemente no hace nada (resultado null).
  await msalInstance.handleRedirectPromise().catch((error) => {
    console.error("[MSAL] Error procesando la redirección:", error);
  });

  // Si ya había una cuenta de una sesión anterior (misma pestaña), la deja
  // activa de una vez para no pedir que la seleccionen otra vez.
  const cuentas = msalInstance.getAllAccounts();
  if (cuentas.length > 0) {
    msalInstance.setActiveAccount(cuentas[0]);
  }

  // Cuando un login termina bien, marca esa cuenta como la activa automáticamente.
  msalInstance.addEventCallback((evento) => {
    if (evento.eventType === EventType.LOGIN_SUCCESS && evento.payload?.account) {
      msalInstance.setActiveAccount(evento.payload.account);
    }
  });

  ReactDOM.createRoot(document.getElementById("root")).render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </React.StrictMode>
  );
}

iniciar();
