// SOURCE: "json" (Opción B, sin login) | "graph" (Opción A, MSAL + Microsoft Graph)
window.CONFIG = {
  SOURCE: "json",

  // Opción B
  JSON_URL: "data.json",

  // Opción A (SPA registrada en Entra ID, permiso delegado Files.Read.All)
  CLIENT_ID: "PEGAR-CLIENT-ID-DE-LA-APP-SPA",
  TENANT_ID: "PEGAR-TENANT-ID",
  REDIRECT_URI: window.location.origin + window.location.pathname,
  OWNER_UPN: "francisco.marambio@pompeyo.cl",
  ITEM_ID: "01CFBOCQQ2ENMHBBNO35ALPFGRXCJ2NMLF",
  TABLE: "Tabla1",

  REFRESH_MIN: 15
};
