# Control de Proyectos ENEL — Mínima Arquitectos

Herramienta web para el seguimiento de los proyectos del contrato ENEL. Cubre cronograma, pendientes de obra, HSE, control financiero (actas, facturación, liquidaciones) y contratistas.

Es una página estática que se publica en **GitHub Pages**. Los datos **no** se guardan en GitHub: viven en **SharePoint (Microsoft 365)** y solo se ven después de iniciar sesión con una cuenta de Mínima.

```
index.html          Estructura de la página
css/style.css       Estilos (identidad Mínima)
js/config.js        ← ÚNICO archivo a editar (conexión Microsoft 365)
js/util.js          Utilidades (fechas, moneda, cruce de nombres)
js/store.js         Almacenamiento: SharePoint vía Microsoft Graph + caché local
js/importer.js      Importador de los Excel actuales
js/app.js           Cálculos, alertas y componentes
js/views.js         Pantallas
assets/             Logo
```

---

## 1. Probarla sin configurar nada (modo local)

1. Abre `index.html` desde GitHub Pages (paso 2) o con cualquier servidor local (`python -m http.server`).
2. Ve a **Datos y conexión** y arrastra los 4 Excel:
   - Seguimiento de liquidaciones
   - Seguimiento semanal (cronograma de ejecución)
   - Cronograma de proyectos
   - Estatus de facturación

En modo local los datos quedan **solo en ese navegador**. Sirve para revisar la importación, no para trabajar en equipo.

## 2. Publicar en GitHub Pages

1. Crea un repositorio, por ejemplo `control-enel`, y sube el contenido de esta carpeta. El `.gitignore` impide subir Excel o JSON por error.
2. En el repositorio entra a **Settings → Pages → Source: Deploy from a branch → main / (root)**.
3. La URL queda así: `https://<usuario-u-organizacion>.github.io/control-enel/`.

> El repositorio puede ser público porque el código no contiene datos ni contraseñas. El *clientId* de Microsoft no es un secreto. Si prefieren un repositorio privado, GitHub Pages exige un plan de pago.

## 3. Conectar con Microsoft 365

Lo hace una vez alguien con acceso de administrador a Microsoft Entra ID (normalmente TI).

### 3.1 Registrar la aplicación
1. Abre <https://entra.microsoft.com> y ve a **Aplicaciones → Registros de aplicaciones → Nuevo registro**.
2. Configura el registro:
   - **Nombre:** Control Proyectos ENEL.
   - **Tipos de cuenta:** *Solo cuentas de este directorio organizativo*.
   - **URI de redirección:** plataforma **Aplicación de página única (SPA)**, con la URL exacta de GitHub Pages, incluida la `/` final. Por ejemplo `https://minima.github.io/control-enel/`.
3. Copia el **Id. de aplicación (cliente)** y el **Id. de directorio (inquilino)**.
4. Agrega los permisos en **Permisos de API → Agregar → Microsoft Graph → Permisos delegados**:
   - `User.Read`
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
5. Pulsa **Conceder consentimiento de administrador**.
6. Recomendado: en **Aplicaciones empresariales → Control Proyectos ENEL → Propiedades** activa *¿Se requiere asignación?* y asigna solo el grupo del equipo ENEL.

### 3.2 Preparar SharePoint
1. Elige un sitio. Por ejemplo `https://minimaarquitectos.sharepoint.com/sites/Proyectos`.
2. En la biblioteca **Documentos** crea la carpeta `Control ENEL`.
   - Si no existe, el archivo de base se crea al primer guardado.
   - Quien no tenga permiso en el sitio no podrá ver los datos, aunque abra la página.

### 3.3 Editar `js/config.js`
```js
msal:    { clientId: 'xxxxxxxx-....', tenantId: 'yyyyyyyy-....', redirectUri: '' },
storage: { type: 'sharepoint', hostname: 'minimaarquitectos.sharepoint.com',
           sitePath: '/sites/Proyectos', library: '', folder: 'Control ENEL',
           file: 'control-enel-db.json' }
```
Sube el cambio a GitHub. Al abrir la página aparecerá **Iniciar sesión con Microsoft**.

### 3.4 Carga inicial
1. Inicia sesión y ve a **Datos y conexión** → importa los 4 Excel.
2. Ve a **Financiero → Actas y facturas** y resuelve los filtros *Sin proyecto* y *Por revisar*. Las actas de ENEL no traen código de proyecto: la herramienta las asigna por similitud de nombre y marca las dudosas.
3. Revisa la lista de nombres unidos que muestra el reporte de importación. Si quedaron proyectos duplicados, usa **Fusionar** en la ficha del proyecto.

Si en el modo local ya había cambios, descarga una copia (JSON) en ese navegador y restáurala después de conectar Microsoft 365.

## 4. Cómo guarda la información

- **Archivo de base:** `Control ENEL/control-enel-db.json` en SharePoint.
  - SharePoint conserva el **historial de versiones**: se puede volver a cualquier versión anterior.
  - Además se guarda **una copia diaria** en `Control ENEL/copias/`.
- **Adjuntos:** las evidencias HSE y los soportes de pendientes se guardan en `Control ENEL/adjuntos/<código del proyecto>/`, con un máximo de 4 MB por archivo.
- **Varias personas a la vez:** cada registro guarda quién lo cambió y cuándo. Si dos personas guardan al tiempo, se unen los cambios registro por registro.
  - Si dos personas editan **el mismo registro**, queda el último guardado.
- **Sin conexión:** los cambios quedan en espera en el navegador y se suben al volver la conexión.
- **Exportar a Excel:** genera un libro con hojas Proyectos, Actividades, Pendientes, HSE, Requisitos HSE, Contratistas, Contratos y Actas. Sirve para informes, Power BI o envíos a ENEL.

No edites el JSON a mano. Para cambios masivos: exporta a Excel, corrige y vuelve a importar, o pide un ajuste al importador.

## 5. Rutina sugerida

| Quién | Cuándo | Qué |
|---|---|---|
| Residentes | Diario / semanal | Pendientes, registros HSE (inspecciones, ATS, permisos), bitácora |
| Coordinador de proyecto | Semanal (antes del comité) | Avance por componente, fechas, estado, observaciones para ENEL |
| Costos / facturación | Por acta | Actas, factura y radicación, estado y valor de la liquidación |
| Compras / SST | Mensual | Contratistas: ARL, planilla, pólizas, contratos y pagos |
| Dirección | Semanal | Tablero, alertas altas y medias |

Una vez la herramienta sea la fuente de verdad, los Excel de seguimiento semanal dejan de mantenerse en paralelo. Si se siguen llenando ambos, los datos se van a desalinear.

## 6. Reglas de cálculo

- **Avance ponderado:**
  - Componentes y pesos: Informe 5 %, Presupuesto 5 %, Preoperacional 10 %, Obra/proyecto 70 %, Liquidación 5 % y Dossier 5 %.
  - Los pesos se cambian en *Datos y conexión*.
- **Obra programada:**
  - Avance lineal entre la fecha de inicio y la de entrega, medido a la fecha de corte.
  - No se calcula si el inicio está "sin confirmar", es decir, si en el Excel el inicio venía de la fecha de corte.
- **Margen:**
  - (valor de liquidación, o el proyectado si aún no hay liquidación) − (contratos de contratistas + otros costos directos).
  - Solo es confiable si todos los costos están registrados.
- **Alertas altas:**
  - Entrega vencida.
  - Obra más de 15 puntos por debajo de lo programado.
  - Permisos de trabajo vencidos sin cerrar.
  - Contratista con ARL o planilla vencida en un contrato vigente.
  - Margen menor al 10 %.
  - Actas conciliadas mayores que una liquidación aprobada (posible nota crédito).
