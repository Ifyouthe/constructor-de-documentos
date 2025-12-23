# Constructor de Documentos Sumate

Microservicio para la generación dinámica de documentos Excel utilizando plantillas almacenadas en Supabase Storage.

## 🚀 Características

- ✅ **Generación de documentos Excel** con ExcelJS
- ✅ **Plantillas dinámicas** almacenadas en Supabase Storage
- ✅ **Sistema de mapeo CSV** para configurar campos
- ✅ **Protección con contraseña** para documentos sensibles
- ✅ **Integración con N8N** via webhooks
- ✅ **Almacenamiento automático** en Supabase Storage
- ✅ **Historial de documentos** con metadata
- ✅ **API REST completa** para gestión de documentos

## 📋 Formatos Soportados

| Formato | Descripción |
|---------|-------------|
| `general` | Documento general Sumate |
| `con_HC` | Scoring con historial crediticio |
| `sin_HC` | Scoring sin historial crediticio |
| `expediente_sumate` | Expediente de cliente |
| `solicitud_credito` | Solicitud de crédito |

## 🔧 Instalación

```bash
npm install
cp .env.example .env
npm run dev
```

## 📡 API Principal

### Webhook de Generación
```http
POST /webhook/generar-documento
```

### Health Check
```http
GET /health
```

### Listar Plantillas
```http
GET /api/plantillas
```

## 🐳 Docker

```bash
docker build -t constructor-documentos-sumate .
docker run -p 3001:3001 --env-file .env constructor-documentos-sumate
```

## 🔧 Configuración

Requiere conexión a Supabase Storage con buckets:
- `plantillas-documentos` - Para plantillas Excel y CSV
- `documentos-generados` - Para documentos creados

Configurar en `.env`:
- SUPABASE_URL
- SUPABASE_ANON_KEY
- N8N_WEBHOOK_URL
- FRASE_SECRETA_EXCEL