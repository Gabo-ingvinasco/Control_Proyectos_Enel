/* ═══════════════════════════════════════════════════════════════════
   CONFIGURACIÓN — es el único archivo que hay que editar.
   Con clientId vacío la herramienta funciona en "modo local":
   los datos quedan solo en el navegador (útil para probar).
   Ver README.md → "Conectar con Microsoft 365".
   ═══════════════════════════════════════════════════════════════════ */
window.APP_CONFIG = {
  appName: 'Control de Proyectos ENEL',

  msal: {
    clientId: '',            // Id. de aplicación (cliente) del registro en Microsoft Entra ID
    tenantId: '',            // Id. de directorio (inquilino) de Mínima
    redirectUri: ''          // vacío = la URL donde está publicada la herramienta
  },

  storage: {
    type: 'sharepoint',                              // 'sharepoint' (recomendado) u 'onedrive' (solo pruebas)
    hostname: 'minimaarquitectos.sharepoint.com',    // dominio SharePoint de la empresa
    sitePath: '/sites/Proyectos',                    // sitio donde vivirá la base
    library: '',                                     // vacío = biblioteca "Documentos"
    folder: 'Control ENEL',                          // carpeta dentro de la biblioteca
    file: 'control-enel-db.json'                     // archivo de la base (no editarlo a mano)
  }
};
