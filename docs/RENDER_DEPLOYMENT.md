# Despliegue en Render Free con Firebase Firestore Spark

Esta guía despliega el backend FastAPI con el runtime Python nativo de Render. No usa
Cloud Run, Cloud Build, Artifact Registry, Secret Manager, Docker Desktop ni procesos
permanentes en la computadora local.

## Arquitectura final

```text
Frontend Vercel
      |
      v
Render Web Service Free (FastAPI + Uvicorn)
      |
      +--> Firebase Cloud Firestore, plan Spark
      |
      +--> OpenAI API
```

La caché conserva este orden:

```text
app/modulos_base.json -> Firestore -> OpenAI -> guardar en Firestore
```

El filesystem de Render es efímero y no se usa para persistencia. Las plantillas PPTX
sí viajan con el código porque son recursos de lectura de la aplicación.

## Datos importantes antes de empezar

Render Free proporciona actualmente una sola instancia con 512 MB de RAM y 0.1 CPU.
Después de 15 minutos sin tráfico el servicio se suspende; la primera petición siguiente
puede tardar aproximadamente un minuto mientras despierta. El plan incluye 750 horas de
instancia mensuales por workspace. Render puede reiniciar una instancia Free y sus
archivos locales se pierden, pero los documentos de Firestore permanecen.

El generador PPTX puede consumir bastante memoria. No se ha modificado esa lógica. Si
una generación concreta supera 512 MB, la solución será reducir ese lote o cambiar de
plan; no se deben agregar workers para intentar resolverlo.

Firestore Spark no solicita información de pago. Su cuota gratuita incluye una base de
datos por proyecto, 1 GiB almacenado, 50 000 lecturas y 20 000 escrituras diarias y
10 GiB de transferencia saliente mensual. No vincules una cuenta de facturación: hacerlo
cambiaría el proyecto al plan Blaze.

## 1. Crear la cuenta y el proyecto de Firebase

1. Entra en <https://console.firebase.google.com/>.
2. Inicia sesión con la cuenta de Google que administrará los datos.
3. Pulsa **Crear un proyecto** o **Agregar proyecto**.
4. Escribe un nombre reconocible, por ejemplo `diplomas-back`.
5. Firebase propondrá un ID único. Anótalo exactamente; en esta guía se llamará
   `FIREBASE_PROJECT_ID`.
6. Si pregunta por Google Analytics, puedes desactivarlo: este backend no lo necesita.
7. Pulsa **Crear proyecto** y espera a que termine.

### Confirmar que continúa en Spark

1. Dentro del proyecto, abre el engranaje junto a **Descripción general del proyecto**.
2. Entra en **Uso y facturación** o **Usage and billing**.
3. Comprueba que el plan mostrado sea **Spark — Sin costo**.
4. No pulses **Modificar plan**, **Upgrade** ni vincules una cuenta de facturación.

## 2. Crear Cloud Firestore

1. En el menú izquierdo entra en **Compilación/Build > Firestore Database**. En la
   interfaz nueva puede aparecer como **Databases & Storage > Firestore**.
2. Pulsa **Crear base de datos**.
3. Selecciona **Standard edition**, no Enterprise.
4. Usa el ID `(default)`.
5. Elige **Production mode**. Este modo bloquea clientes web no autorizados, pero permite
   el acceso del backend autenticado con su cuenta de servicio.
6. Elige `us-east4 (Northern Virginia)` si está disponible. Render se configurará en
   Virginia, por lo que esta ubicación reduce la comunicación entre backend y Firestore.
   La ubicación de Firestore no puede cambiarse después.
7. Pulsa **Crear** y espera a que aparezca la pestaña **Datos**.
8. Vuelve a **Uso y facturación** y confirma nuevamente **Spark**.

No crees manualmente la colección: el script de migración creará `modulos_cache`.

## 3. Obtener la credencial del servidor

La alternativa elegida es un **Secret File de Render**. Es más sencilla que copiar JSON
multilínea dentro de una variable y funciona con el mecanismo estándar de Google
`GOOGLE_APPLICATION_CREDENTIALS`.

1. En Firebase abre el engranaje y selecciona **Configuración del proyecto**.
2. Abre la pestaña **Cuentas de servicio / Service accounts**.
3. Selecciona **Firebase Admin SDK**.
4. Pulsa **Generar nueva clave privada** y confirma **Generar clave**.
5. El navegador descargará un JSON. No lo abras en editores o servicios que lo sincronicen
   públicamente y no lo envíes por correo o chat.
6. Guárdalo fuera del repositorio. Ejemplo seguro local:

```text
C:\Users\cueva\secrets\diplomas-back\firebase-service-account.json
```

7. Abre el JSON localmente y comprueba que `project_id` coincide con el proyecto Firebase.
8. No subas este archivo a Git. El `.gitignore` ya bloquea nombres habituales, pero la
   ubicación fuera del repositorio es la protección principal.

Si una clave se publica accidentalmente, elimínala inmediatamente en **Cuentas de
servicio**, genera otra y reemplaza el Secret File de Render.

## 4. Preparar Python local

El proyecto fija Python 3.13 mediante `.python-version`. Render admite ese formato y usa
la versión de parche 3.13 más reciente disponible.

Si necesitas recrear el entorno local:

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## 5. Validar y migrar los 466 registros

Desde la raíz del repositorio, valida primero sin conectarte a Firebase:

```powershell
python scripts\migrate_cache_to_firestore.py --dry-run
```

La salida esperada actualmente es:

```text
Registros encontrados: 466
Registros válidos: 466
Errores de validación: 0
DRY RUN finalizado: no se realizó ninguna conexión ni escritura en Firestore.
```

Después ejecuta la migración real indicando la credencial guardada fuera del repositorio:

```powershell
$FIREBASE_PROJECT_ID = "tu-id-real-de-firebase"
$FIREBASE_KEY_PATH = "C:\Users\cueva\secrets\diplomas-back\firebase-service-account.json"

python scripts\migrate_cache_to_firestore.py `
  --project $FIREBASE_PROJECT_ID `
  --credentials-file $FIREBASE_KEY_PATH
```

El script usa la misma normalización que la aplicación y `set(..., merge=True)`. Puedes
repetirlo: actualiza cada ID existente y no crea duplicados.

### Verificar la migración

1. Regresa a **Firebase Console > Firestore Database > Datos**.
2. Debe aparecer la colección `modulos_cache`.
3. Abre varios documentos. Sus IDs tendrán formas como:
   `8_MODULOS::derecho-penal`.
4. Comprueba que cada uno incluya `grupo`, `tipo_referencia`, `tema_original` y el arreglo
   `modulos`; los registros históricos también conservarán sus metadatos de fecha.
5. La consola debe mostrar 466 documentos si ninguno fue agregado posteriormente.

No elimines todavía `app/data/modulos_cache.json`; mantenlo como respaldo local privado.

## 6. Subir el código a un repositorio Git

Render necesita acceder a GitHub, GitLab o Bitbucket.

1. Ejecuta `git status` y confirma que ningún JSON de cuenta de servicio aparece listado.
2. Confirma que `.env` tampoco aparece.
3. Revisa y confirma los cambios del proyecto en tu repositorio remoto.
4. Comprueba en la web del proveedor Git que no existe ningún archivo con nombres como
   `firebase-service-account.json` o `firebase-adminsdk...json`.

## 7. Crear la cuenta de Render

1. Entra en <https://dashboard.render.com/register>.
2. Regístrate preferiblemente con el mismo proveedor Git donde está este repositorio.
3. Autoriza a Render únicamente para los repositorios necesarios cuando el proveedor
   permita elegirlos.
4. Entra en el Dashboard. No es necesario crear Postgres ni Key Value.

## 8. Crear el Web Service Free

El archivo `render.yaml` documenta una configuración reproducible. Para el primer
despliegue de un principiante se recomienda crear el Web Service manualmente, porque así
puedes agregar el Secret File antes de probar la API.

1. En Render pulsa **New + > Web Service**.
2. Selecciona **Git Provider** y conecta GitHub/GitLab/Bitbucket si aún no está conectado.
3. Busca este repositorio y pulsa **Connect**.
4. Configura:

| Campo | Valor exacto |
|---|---|
| Name | `diplomas-back` o un nombre disponible |
| Language/Runtime | `Python 3` |
| Branch | la rama donde confirmaste estos cambios |
| Region | `Virginia (US East)` |
| Root Directory | vacío, si este archivo está en la raíz del repo |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn app.main:app --host 0.0.0.0 --port $PORT` |
| Instance Type | `Free` |
| Health Check Path | `/health` |

5. No agregues workers ni antepongas Gunicorn. El comando inicia un único Uvicorn.
6. Render leerá `.python-version`. Después verifica en el log de construcción que aparezca
   `Python 3.13.x`.

## 9. Configurar variables privadas y Firestore en Render

Antes de hacer pruebas funcionales, abre la sección **Environment** o **Advanced** del
formulario/servicio y agrega estas variables:

| Key | Value | Secreto |
|---|---|---|
| `OPENAI_API_KEY` | tu clave real de OpenAI | Sí |
| `FIRESTORE_PROJECT_ID` | el ID exacto del proyecto Firebase | No es clave, pero puede mantenerse privado |
| `FIRESTORE_DATABASE_ID` | `(default)` | No |
| `FIRESTORE_COLLECTION` | `modulos_cache` | No |
| `GOOGLE_APPLICATION_CREDENTIALS` | `/etc/secrets/firebase-service-account.json` | No contiene la clave, solo la ruta |

No pegues la clave de OpenAI en `render.yaml` ni en Git.

### Agregar el Secret File

1. En la misma página busca **Secret Files**.
2. Pulsa **Add Secret File**.
3. En **Filename** escribe exactamente:

```text
firebase-service-account.json
```

4. Abre el JSON descargado de Firebase con un editor local, copia todo su contenido —desde
   la primera `{` hasta la última `}`— y pégalo en **Contents**.
5. Guarda. Render lo montará en:

```text
/etc/secrets/firebase-service-account.json
```

6. Comprueba que coincide exactamente con `GOOGLE_APPLICATION_CREDENTIALS`.
7. Nunca imprimas el contenido del archivo en los logs.

Si utilizas **New > Blueprint** en lugar de la creación manual, Render leerá
`render.yaml` y solicitará `OPENAI_API_KEY` y `FIRESTORE_PROJECT_ID`. El Secret File no
está versionado en el Blueprint: agrégalo manualmente en **Environment > Secret Files**
inmediatamente después de crear el servicio y antes de probar `/api/diplomas`.

## 10. Desplegar

1. Revisa una vez más que el plan seleccionado sea **Free**.
2. Pulsa **Create Web Service** o **Deploy latest commit**.
3. Observa los logs. El build debe instalar `requirements.txt` y el proceso debe iniciar
   con Uvicorn.
4. El health check `/health` debe pasar antes de que Render marque el deploy como **Live**.
5. Copia la URL pública indicada por Render, por ejemplo:

```text
https://diplomas-back.onrender.com
```

No necesitas dejar esta computadora encendida después del despliegue.

## 11. Probar `/health`

La primera petición puede tardar por el cold start del plan Free:

```powershell
$RENDER_URL = "https://tu-servicio.onrender.com"
Invoke-RestMethod "$RENDER_URL/health"
```

Resultado esperado:

```json
{"status":"ok"}
```

## 12. Probar `/api/diplomas`

La primera prueba usa un tema del JSON base y no debería consultar Firestore ni OpenAI:

```powershell
$TEST_BODY = @{
  items = @(
    @{
      modeloCertificado = "INSTITUTO"
      tipoModelo = "DIPLOMADO"
      nombres = "Persona"
      apellidos = "De Prueba"
      temaDiplomado = "Inteligencia Artificial Aplicada a la Educación"
      fechaInicio = "2026-01-01"
      fechaFin = "2026-02-01"
      horasAcademicas = 120
      creditosAcademicos = 8
      folioNumero = "PRUEBA-001"
      fechaEmision = "2026-02-02"
      codigoEstudiante = "TEST001"
      ciudad = "Lima"
    }
  )
} | ConvertTo-Json -Depth 5

Invoke-WebRequest `
  -Method Post `
  -Uri "$RENDER_URL/api/diplomas" `
  -ContentType "application/json" `
  -Body $TEST_BODY `
  -OutFile ".\diploma-prueba-render.pptx"
```

Confirma que se descargó un PPTX y que Render no informó falta de memoria.

## 13. Comprobar la reutilización de Firestore

Usa un tema migrado que no está en el JSON base, por ejemplo `Derecho Penal` con tipo
`DIPLOMADO`:

1. Repite el cuerpo anterior cambiando `temaDiplomado` por `Derecho Penal`.
2. Conserva `tipoModelo = "DIPLOMADO"` para buscar el documento
   `8_MODULOS::derecho-penal`.
3. Envía la petición dos veces.
4. En **Render > Logs** debe aparecer:

```text
Caché de módulos encontrada en Firestore: cache_key=8_MODULOS::derecho-penal
```

5. Comprueba también que el documento continúa en Firebase.
6. En el panel de uso de OpenAI verifica que estas repeticiones no hayan creado llamadas
   para generar módulos.

Para probar una escritura nueva, utiliza posteriormente un tema verdaderamente nuevo,
realiza una solicitud y confirma que Firebase crea el documento normalizado. No repitas
muchos temas de prueba para no consumir OpenAI innecesariamente.

## 14. Conectar Vercel

El origen productivo existente ya está permitido:

```text
https://ipdefrontendcertificados.vercel.app
```

En Vercel:

1. Abre el proyecto del frontend.
2. Entra en **Settings > Environment Variables**.
3. Identifica la variable que actualmente contiene la URL del backend.
4. Sustituye su valor por la URL `https://...onrender.com`, sin añadir rutas como
   `/health` o `/api/diplomas` si el frontend ya las concatena.
5. Aplica el valor al entorno **Production**.
6. Vuelve a desplegar el frontend.
7. Prueba desde el dominio productivo y revisa la consola del navegador por si aparece
   algún error CORS.

CORS no requiere cambios mientras el dominio de Vercel sea exactamente el indicado.

## 15. Límites y seguridad

- Render Free es apropiado para pruebas, proyectos personales y tráfico moderado, no para
  disponibilidad garantizada.
- Tras 15 minutos sin tráfico habrá un cold start. No crees monitores artificiales para
  mantenerlo despierto: consumirían horas gratuitas y pueden incumplir el propósito del plan.
- Si no agregas un medio de pago y agotas límites de Render, el servicio puede suspenderse
  hasta el siguiente periodo en vez de generar un cobro.
- Firestore Spark detiene el servicio al superar las cuotas gratuitas; no hay consumo
  adicional facturable mientras el proyecto continúe en Spark.
- OpenAI se factura de forma independiente de Render y Firebase.
- Este backend es público. Antes de darle tráfico importante conviene agregar autenticación
  o rate limiting en una tarea separada.
