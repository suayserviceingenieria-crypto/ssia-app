// ============================================================================
// PUNTO DE ENTRADA — S&S IA
// ============================================================================
// Envuelve toda la app en MsalProvider para que cualquier componente pueda
// usar los hooks de @azure/msal-react (useMsal, useIsAuthenticated, etc.)
// para la conexión con Microsoft 365 / OneDrive. Esto es independiente del
// login por usuario/clave (SHA-256) que ya maneja App.jsx internamente.
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

  // Si ya había una cuenta de una sesión anterior (misma pestaña), la deja
  // activa de una vez para no pedir que la seleccionen otra vez.
  const cuentas = msalInstance.getAllAccounts();
  if (cuentas.length > 0) {
    msalInstance.setActiveAccount(cuentas[0]);
  }

  // Cuando un login (popup o redirect) termina bien, marca esa cuenta como
  // la activa automáticamente.
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
