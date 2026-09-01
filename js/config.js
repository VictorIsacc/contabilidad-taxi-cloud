// CONFIGURACIÓN DE CONTABILIDAD TAXI CLOUD
// El Client ID se añadirá cuando registremos la aplicación en Microsoft.
export const CLOUD_CONFIG = {
  microsoftClientId: "PENDIENTE_DE_CONFIGURAR",
  tenant: "consumers",
  redirectUri: window.location.origin + window.location.pathname.replace(/index\.html$/, ""),
  scopes: ["User.Read", "Files.ReadWrite"],
  workbook: {
    itemId: localStorage.getItem("taxiWorkbookItemId") || "",
    name: localStorage.getItem("taxiWorkbookName") || ""
  }
};
