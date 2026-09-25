# Confirmación de descarga en Angular

Los cambios ya se aplicaron localmente en `C:\Benja\Instituto\diplomas-front`; no es necesario volver a aplicar el parche en ese repositorio. El archivo `.patch` conserva los cambios para revisión o para aplicarlos sobre otra copia de la versión original del frontend.

La confirmación se envía después de iniciar la descarga del Blob recibido. Solo la confirmación tiene reintentos (1 y 2 segundos); la generación nunca se reintenta automáticamente. Requiere desplegar el backend con el header CORS expuesto `X-Generation-Id` y el endpoint `POST /api/diplomas/confirmar/{generation_id}`.

Para verificar los cambios locales desde PowerShell:

```powershell
Set-Location C:\Benja\Instituto\diplomas-front
npm test -- --watch=false --include="src/app/**/*.generation-log.spec.ts"
npm run build
```

El parche agrega pruebas focalizadas nuevas y no modifica las pruebas heredadas. Verifica URL y tiempo máximo de la confirmación, reintentos acotados y errores permanentes, ausencia de reintentos de generación, orden descarga/confirmación, respuesta sin Blob, compatibilidad sin header y fallo de confirmación sin repetir descarga.

Las pruebas usan datos sintéticos y solicitudes HTTP simuladas. No generan registros reales en Firestore o Google Sheets.
