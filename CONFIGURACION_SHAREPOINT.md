# Configuración de SharePoint para Ejecución Presupuestaria

El módulo busca automáticamente la carpeta mensual más reciente dentro de
`PLANIFICACIÓN TÉCNICA/AÑO {year}/REPORTE ESIGEF` y selecciona el Excel con la
fecha más alta cuyo nombre siga el patrón `Nacional al dd.mm.aa.xlsx`.

## Variables de entorno de la aplicación publicada

Configure estos valores como secretos en el entorno donde se publique Dash:

```text
MS_TENANT_ID=<identificador del tenant institucional>
MS_CLIENT_ID=<identificador de la aplicación registrada>
MS_CLIENT_SECRET=<secreto de la aplicación>
SHAREPOINT_USER_UPN=planificacion_tecnica@educacion.gob.ec
ESIGEF_ROOT_PATH=PLANIFICACIÓN TÉCNICA/AÑO {year}/REPORTE ESIGEF
```

No guarde `MS_CLIENT_SECRET` dentro de `app.py`, GitHub ni ningún Excel.

## Permiso requerido

La aplicación registrada en Microsoft Entra ID necesita permiso de aplicación
de solo lectura para Microsoft Graph y consentimiento administrativo. Para
restringir el alcance, el administrador institucional puede preferir
`Sites.Selected`; si no está configurado ese esquema, puede habilitar
`Files.Read.All` para esta aplicación de consulta.

La aplicación vuelve a comprobar SharePoint cada 15 minutos. Si todavía no
existen credenciales, también admite como respaldo un archivo local con nombre
`Nacional al dd.mm.aa.xlsx` dentro de la carpeta del proyecto.
