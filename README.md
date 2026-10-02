# Constructor de Documentos Sumate

Microservicio que llena plantillas Excel (ExcelJS) y Word (docxtemplater) con los datos de un prospecto. Lo usa Dari (`EXCEL_API_URL`) para las fichas de cada prospecto que se envía a sucursal.

Las plantillas y sus mapfields se leen de un bucket **S3 compatible** (MinIO en el 207) o de un directorio local. Ya no depende de Supabase.

## Formatos

| `formato` | Plantilla | Hoja |
|---|---|---|
| `general`, `identificacion_cliente`, `formato_editable` | `Formato_Editable_Listo.xlsx` | Ficha de identificación |
| `con_HC` | `SCORING_CON_HC.xlsx` | Scoring del Cliente |
| `sin_HC` | `SCORING_SIN_HC.xlsx` | Scoring del Cliente |
| `seguimiento` | `seguimiento del credito (1).xlsx` | Hoja1 |
| `evaluacion_economica` | `Evaluacion_Economica_con_Etiquetas.xlsx` | EE Simple 2025 |
| `scoring_etiquetas` | `Scoring con Etiquetas.xlsx` | Scoring |
| `obligado_solidario`, `visita_domiciliaria`, `ficha_aval` | `.docx` correspondiente | — |

## API

Todas las rutas, salvo `GET /health`, exigen la cabecera `X-API-Key` cuando `CONSTRUCTOR_API_KEY` tiene valor (si no, 401).

| Método y ruta | Uso |
|---|---|
| `POST /webhook/generar-documento` | Devuelve el `.xlsx`/`.docx` (lo que usa Dari) |
| `POST /webhook/generar-multiples-documentos` | JSON con los documentos en base64 |
| `POST /webhook/generar-multiples-documentos-zip` | ZIP (o el archivo si es uno) |
| `POST /api/generar-documento` | JSON con base64 (pruebas) |
| `POST /api/generar-multiples-documentos` | JSON con base64 |
| `GET /api/plantillas` | Lista de plantillas |
| `GET /health` | Sin clave; 503 si no se pueden listar las plantillas |

Se retiraron `/api/documentos`, `/api/descargar/:id` y `/api/stats`: leían la tabla `documentos_generados_sumate` de Supabase, que siempre estuvo vacía.

```bash
curl -X POST http://sumate-constructor:3003/webhook/generar-documento \
  -H "Content-Type: application/json" -H "X-API-Key: $CONSTRUCTOR_API_KEY" \
  -d '{"formato":"con_HC","buro.BC_score":"712","codigo_de_prospecto":"PRUEBA-1"}' --output ficha.xlsx
```

## Variables de entorno

Ver `.env.example`.

| Variable | Obligatoria | Por defecto | Para qué |
|---|---|---|---|
| `CONSTRUCTOR_API_KEY` | sí en servidores | vacía (abierto) | Secreto compartido de la cabecera `X-API-Key` |
| `STORAGE_DRIVER` | no | `s3` | `s3` (MinIO/S3) o `local` (directorio) |
| `MINIO_ENDPOINT` | con `s3` | — | p. ej. `http://sumate-minio:9000` |
| `MINIO_ACCESS_KEY`, `MINIO_SECRET_KEY` | con `s3` | — | Usuario de **solo lectura** sobre el bucket de plantillas |
| `MINIO_REGION` | no | `us-east-1` | |
| `MINIO_FORCE_PATH_STYLE` | no | `true` | `false` para S3 de AWS con virtual-host |
| `BUCKET_PLANTILLAS` | no | `plantillas-constructor` | Bucket con los 18 archivos |
| `PLANTILLAS_DIR` | con `local` | `./plantillas` | Directorio con los 18 archivos |
| `BUCKET_GENERADOS` / `GENERADOS_DIR` | no | vacía | Solo para `saveToStorage: true` (nadie lo usa hoy) |
| `MINIO_URL_PUBLICA` | no | `MINIO_ENDPOINT` | Base de la URL que se devuelve con `saveToStorage` |
| `FRASE_SECRETA_EXCEL` | no | — | Solo formato `expediente_sumate` (contraseña de la hoja) |
| `MAX_BODY_SIZE` | no | `1mb` | Límite del cuerpo JSON (413 si se pasa) |
| `RATE_LIMIT_WINDOW_MS`, `RATE_LIMIT_MAX_REQUESTS` | no | `900000`, `50` | Límite en `/api/*` |
| `PORT` | no | `3003` | |

## Plantillas: copia desde Supabase

`scripts/copiar-plantillas.js` lee (solo lectura: `list` y `download`) el bucket `documentossumate` de supabase-evolve, sube cada archivo al bucket S3 y/o a un directorio, y lo verifica por sha256 releyéndolo del destino:

```bash
SUPABASE_URL=... SUPABASE_KEY=<anon> SUPABASE_BUCKET=documentossumate \
MINIO_ENDPOINT=http://127.0.0.1:9000 MINIO_ACCESS_KEY=<admin> MINIO_SECRET_KEY=... \
BUCKET_PLANTILLAS=plantillas-constructor \
node scripts/copiar-plantillas.js --s3 [--dir ./plantillas]
# o, desde una carpeta ya descargada:
node scripts/copiar-plantillas.js --desde-dir ./plantillas --s3
```

Las plantillas son documentos internos de Súmate: **no se suben a este repositorio** (es público).

## Pruebas

```bash
npm test                                      # unitarias + servidor con plantilla sintética
TEST_MINIO_ENDPOINT=http://127.0.0.1:9000 TEST_MINIO_ACCESS_KEY=... TEST_MINIO_SECRET_KEY=... npm test   # + MinIO real
node scripts/comparar-documentos.js a.xlsx b.xlsx [...]   # equivalencia por partes del ZIP y celda a celda
```

## Despliegue en el 207

- Rama `refactor/207` (no `main`: `main` se autodespliega en el Dokploy del 195 con Supabase).
- `docker compose -p sumate-constructor -f docker-compose.yml -f docker-compose.207.yml up -d --build`, con `.env` 600.
- Contenedor `sumate-constructor`, puerto **3003** sin publicar, solo en la red externa `sumate_servicios` (tiene que existir antes). El MinIO de `MINIO_ENDPOINT` tiene que estar en una red común.
- Clientes: `EXCEL_API_URL=http://sumate-constructor:3003/webhook/generar-documento` y la cabecera `X-API-Key`.

## Docker local

```bash
docker build -t constructor-documentos-sumate .
docker run -p 3003:3003 --env-file .env constructor-documentos-sumate
```
