import { CLOUD_CONFIG } from "./config.js";

export function microsoftConfigured(){
  return CLOUD_CONFIG.microsoftClientId &&
         CLOUD_CONFIG.microsoftClientId !== "PENDIENTE_DE_CONFIGURAR";
}

export async function signInMicrosoft(){
  if(!microsoftConfigured()){
    throw new Error("Todavía falta registrar la aplicación en Microsoft y añadir el Client ID.");
  }
  // En el siguiente paso integraremos MSAL Browser aquí.
  throw new Error("Client ID detectado, pero MSAL todavía no está activado en esta versión base.");
}

export async function getAccessToken(){
  throw new Error("Microsoft todavía no está conectado.");
}

export function signOutMicrosoft(){
  sessionStorage.removeItem("taxiAccessToken");
}
