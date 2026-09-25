# Control de PPTX generados y descargas iniciadas

El backend guarda **un documento por PPTX**, con todos sus certificados dentro de
`items`, en la colección `diplomas_generados`. Google Sheets conserva los 17 campos
propuestos, con **una fila por certificado**. Todas las filas de un PPTX tienen el
mismo `generation_id`: contar filas contaría certificados, no descargas.

## Comportamiento implementado

1. Generar y guardar el PPTX en memoria, incluidas las dos rutas de plantillas.
2. Crear su registro `GENERADO` en Firestore y devolver `X-Generation-Id` junto al archivo.
3. Después de enviar el archivo, intentar sincronizar Sheets.
4. Angular recibe el Blob, inicia la descarga y confirma el ID con
   `POST /api/diplomas/confirmar/{generation_id}`.
5. Firestore conserva la primera confirmación y programa otra sincronización.

`CONFIRMADO_FRONTEND` significa archivo recibido por Angular y descarga iniciada;
el navegador no permite comprobar el guardado físico en disco. La fecha de registro
está en B y la fecha de confirmación en Q; ambas incluyen hora de Perú y `-05:00`.
Para contar descargas por fecha/hora usa Q. Firestore también conserva timestamps UTC,
nombre del archivo y `cantidad_certificados`.

Reintentar una confirmación conserva ID y hora. Volver a generar voluntariamente
otro PPTX crea otro ID, aun si contiene los mismos participantes y temas.

Si Firestore no puede guardar el registro inicial, la respuesta es 503 y no entrega
un archivo sin registro. Si Sheets falla, el archivo sí se entrega y el registro
principal permanece en Firestore. Una confirmación fallida se reintenta dos veces
desde Angular, sin volver a generar ni descargar el PPTX. Si falla definitivamente,
el registro permanece `GENERADO`; no se inventa una descarga confirmada.

## Configuración pendiente en Google y Render

1. Usar el proyecto Firebase existente y habilitar **Google Sheets API** en ese
   mismo proyecto de Google Cloud.
2. En el documento compartido, usar la pestaña **Registros**.
3. Pegar esta línea en A1 (campos separados por tabulaciones):

```text
generation_id	fecha_registro	estado	modelo_certificado	tipo_modelo	nombres	apellidos	tema	fecha_inicio	fecha_fin	horas	creditos	folio	fecha_emision	codigo_estudiante	ciudad	confirmado_frontend_at
```

4. Compartir el documento como **Editor** con el `client_email` de la cuenta de
   servicio existente. La hoja puede seguir siendo privada.
5. En Render, agregar:

```text
GOOGLE_SHEET_ID=1WcKFW64W-4Kwm9XFAOQkY3e3ldXv4_Muc-AJ2-QFiq4
```

Conservar `GOOGLE_APPLICATION_CREDENTIALS=/etc/secrets/firebase-service-account.json`
y el Secret File existente. No hace falta una nueva clave. No activar un plan de pago
para esta integración. Si no hay Sheet ID, Firestore funciona y deja `NO_CONFIGURADO`.

La dependencia `google-api-python-client==2.200.0` ya está en requirements.txt.
Desplegar primero el backend y después el frontend actualizado. El frontend conserva
compatibilidad con un backend anterior sin el nuevo header, pero ese caso no registra
confirmaciones. El código local no cambia automáticamente los despliegues.

## Contar PPTX por fecha y hora en Sheets

Puedes añadir R1=`fecha_descarga` y S1=`hora_descarga`; el backend solo escribe A:Q.
En R2 y S2 respectivamente (arrastrar hacia abajo):

```text
=IF(Q2="","",LEFT(Q2,10))
=IF(Q2="","",MID(Q2,12,2)&":00")
```

En una pestaña **Resumen**, poner en A2 una fecha como texto `2026-09-24` y en B2
una hora como texto `10:00`. PPTX confirmados de ese día:

```text
=COUNTUNIQUEIFS(Registros!A2:A,Registros!C2:C,"CONFIRMADO_FRONTEND",Registros!R2:R,A2)
```

PPTX confirmados de ese día y esa hora:

```text
=COUNTUNIQUEIFS(Registros!A2:A,Registros!C2:C,"CONFIRMADO_FRONTEND",Registros!R2:R,A2,Registros!S2:S,B2)
```

Las fórmulas usan nombres ingleses y comas; según la configuración de Sheets,
habilita nombres de funciones en inglés y sustituye comas por punto y coma.
Los participantes, temas y modelos están en cada fila del detalle; filtrar por ID
muestra todos los certificados incluidos en ese archivo.

## Verificación y recuperación

En local, configurar `.env` usando `.env.example`, con la ruta local real de la
credencial (no la ruta de Render). No publicar el JSON de credenciales.

Comprobación **solo lectura**, sin crear registros ni archivos:

```powershell
.\.venv\Scripts\python.exe scripts\sync_generation_logs.py --check
```

Esto comprueba acceso de lectura y encabezados; el permiso Editor debe verificarse
en Compartir o mediante una generación real. Para recuperar hasta 10 generaciones
pendientes en Sheets:

```powershell
.\.venv\Scripts\python.exe scripts\sync_generation_logs.py --limit 10
```

Repetir hasta que `revisados` sea cero. Si hay un error persistente, corregir primero
permisos/API/encabezados; no repetir en bucle. El script escribe en Sheets y actualiza
el estado de sincronización en Firestore, sin generar PPTX ni llamar a OpenAI.

La sincronización busca el ID antes de añadir filas. Un bloqueo temporal en Firestore
evita envíos simultáneos del mismo lote desde el servidor y el script. Los bloqueos
expiran tras 5 minutos si Render se reinicia. Los errores quedan como `ERROR`; los
reinicios durante una tarea pueden dejar `PENDIENTE` o `SINCRONIZANDO` recuperables.
Render Free no ejecuta una cola permanente: se intenta al generar/confirmar; para
pendientes que sigan sin sincronizar, ejecutar el comando anterior.

Usar vistas de filtro para consultar. Evitar insertar, borrar u ordenar físicamente
filas mientras se está sincronizando, ya que Sheets no ofrece una transacción conjunta
con Firestore. Si el número de filas existentes para un lote es inconsistente, se
marca error para revisión en lugar de añadir duplicados. La garantía no es una
transacción distribuida: el conteo autorizado es Firestore y el conteo en Sheets debe
usar IDs únicos.

Diagnóstico habitual: HTTP 403 = API deshabilitada o permisos insuficientes; 404 =
ID incorrecto o documento inaccesible; 400 = pestaña/rango inválido; `ValueError` =
encabezados o cantidad de filas del lote inconsistentes. No se imprimen credenciales
ni los datos de participantes en logs de esta integración.

Prueba de aceptación tras desplegar: generar un PPTX de 3 certificados, comprobar
un documento en Firestore, 3 filas con el mismo ID y `CONFIRMADO_FRONTEND` en Sheets;
el total de PPTX debe aumentar **1**. Repetir una confirmación no debe sumar ni alterar
la primera hora. No se puede reconstruir el histórico anterior a la activación.

Referencias: [append de Sheets](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets.values/append),
[transporte por hilo](https://googleapis.github.io/google-api-python-client/docs/thread_safety.html),
[transacciones Firestore](https://firebase.google.com/docs/firestore/manage-data/transactions).
