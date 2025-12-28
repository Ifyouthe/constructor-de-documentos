// =============================================
// CONSTRUCTOR DE DOCUMENTOS SUMATE - Microservicio
// Generación dinámica de documentos Excel con plantillas
// =============================================

const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
require('dotenv').config();

// Importar configuraciones
const { checkSupabaseConnection, storageUtils, documentUtils } = require('./src/config/supabase');

// Importar servicios
const ExcelService = require('./src/services/excel/excelService');
const WordService = require('./src/services/word/wordService');

// Instanciar servicios
const excelService = new ExcelService();
const wordService = new WordService();

const app = express();
const PORT = process.env.PORT || 3003;

// Configurar trust proxy para Traefik
app.set('trust proxy', true);

console.log('========================================');
console.log('🏗️  INICIANDO CONSTRUCTOR DE DOCUMENTOS SUMATE');
console.log('📋 Microservicio de Generación de Documentos Excel');
console.log(`🔍 Puerto: ${PORT}`);
console.log(`🌍 Entorno: ${process.env.NODE_ENV || 'development'}`);
console.log('========================================');

// =============================================
// MIDDLEWARES
// =============================================

// Seguridad
app.use(helmet({
  contentSecurityPolicy: false
}));

// CORS
app.use(cors({
  origin: ['https://sumate.evolvedigital.cloud', 'http://localhost:3000', 'http://localhost:3001'],
  credentials: true
}));

// Rate limiting
const limiter = rateLimit({
  windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000, // 15 minutos
  max: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS) || 50,
  message: {
    error: 'Demasiadas solicitudes desde esta IP',
    retryAfter: 'Intente nuevamente en 15 minutos'
  },
  standardHeaders: true,
  legacyHeaders: false
});

app.use('/api/', limiter);

// Body parsing
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// =============================================
// ENDPOINTS PRINCIPALES
// =============================================

/**
 * Health Check
 */
app.get('/health', async (req, res) => {
  try {
    const supabaseCheck = await checkSupabaseConnection();

    // Verificar storage de plantillas
    const templatesCheck = await storageUtils.listTemplates();

    const healthStatus = {
      status: 'ok',
      timestamp: new Date().toISOString(),
      service: 'constructor-de-documentos-sumate',
      version: '1.0.0',
      environment: process.env.NODE_ENV,
      port: PORT,
      uptime: process.uptime(),
      memory: {
        used: Math.round(process.memoryUsage().heapUsed / 1024 / 1024),
        total: Math.round(process.memoryUsage().heapTotal / 1024 / 1024)
      },
      checks: {
        supabase: supabaseCheck.success ? 'healthy' : 'error',
        storage: templatesCheck.success ? 'healthy' : 'error',
        templatesCount: templatesCheck.templates ? templatesCheck.templates.length : 0
      },
      errors: []
    };

    if (!supabaseCheck.success) {
      healthStatus.errors.push(`Supabase: ${supabaseCheck.error}`);
    }

    if (!templatesCheck.success) {
      healthStatus.errors.push(`Storage: ${templatesCheck.error}`);
    }

    const hasErrors = !supabaseCheck.success || !templatesCheck.success;

    res.status(hasErrors ? 503 : 200).json(healthStatus);

  } catch (error) {
    console.error('[HEALTH] Error en health check:', error.message);
    res.status(500).json({
      status: 'error',
      timestamp: new Date().toISOString(),
      service: 'constructor-de-documentos-sumate',
      error: 'Error interno del servidor'
    });
  }
});

/**
 * Generar documento desde webhook (endpoint principal)
 */
app.post('/webhook/generar-documento', async (req, res) => {
  try {
    console.log('[WEBHOOK] 📨 Solicitud de generación de documento recibida');
    console.log('[WEBHOOK] 📊 Datos:', JSON.stringify(req.body, null, 2));

    // Detectar tipo de documento basado en formato o tipo especificado
    const formato = req.body.formato || req.body.template || 'general';
    const documentType = req.body.type ||
      (formato === 'obligado_solidario' || formato === 'obligado' || formato === 'ficha_obligado' ||
       formato === 'visita_domiciliaria' || formato === 'ficha_aval' ? 'word' : 'excel');

    let result;

    if (documentType === 'word') {
      // Configurar template correcto para Word
      const dataConTemplate = { ...req.body };
      if (formato === 'obligado_solidario' || formato === 'obligado' || formato === 'ficha_obligado') {
        dataConTemplate.template = 'Fichadeidentificaciondelobligadosolidarioconetiquetas.doc';
      } else if (formato === 'visita_domiciliaria') {
        dataConTemplate.template = 'Visita domiciliaria con etiquetas.doc';
      } else if (formato === 'ficha_aval') {
        dataConTemplate.template = 'Ficha de identificación del aval con etiquetas.doc';
      }

      result = await wordService.processWebhookData(dataConTemplate);
    } else {
      result = await excelService.processWebhookData(req.body);
    }

    if (!result.success) {
      console.error('[WEBHOOK] ❌ Error procesando:', result.error);
      return res.status(400).json({
        success: false,
        error: result.error,
        timestamp: new Date().toISOString()
      });
    }

    console.log('[WEBHOOK] ✅ Documento generado exitosamente:', result.fileName);

    // Devolver siempre el archivo construido directamente
    const mimeType = documentType === 'word'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    res.setHeader('Content-Type', mimeType);
    res.setHeader('Content-Disposition', `attachment; filename="${result.fileName}"`);
    res.send(result.buffer);

  } catch (error) {
    console.error('[WEBHOOK] ❌ Error interno:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error interno del servidor',
      timestamp: new Date().toISOString()
    });
  }
});

/**
 * Generar documento directo (para testing)
 */
app.post('/api/generar-documento', async (req, res) => {
  try {
    const { data, formato = 'general' } = req.body;

    if (!data) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren datos para generar el documento'
      });
    }

    console.log(`[API] 📋 Generando documento formato: ${formato}`);

    let result;

    // Determinar si usar Word o Excel basado en el formato
    if (formato === 'obligado_solidario' || formato === 'obligado' || formato === 'ficha_obligado' ||
        formato === 'visita_domiciliaria' || formato === 'ficha_aval') {
      // Usar servicio Word para estos formatos
      let templateName;
      switch(formato) {
        case 'obligado_solidario':
        case 'obligado':
        case 'ficha_obligado':
          templateName = 'Fichadeidentificaciondelobligadosolidarioconetiquetas.doc';
          break;
        case 'visita_domiciliaria':
          templateName = 'Visita domiciliaria con etiquetas.doc';
          break;
        case 'ficha_aval':
          templateName = 'Ficha de identificación del aval con etiquetas.doc';
          break;
        default:
          templateName = formato;
      }

      result = await wordService.generateWord(data, templateName);
      result.formato = formato; // Mantener el formato original para referencia
    } else {
      // Usar servicio Excel para los demás formatos
      result = await excelService.generateExcel(data, formato);
    }

    if (!result.success) {
      return res.status(400).json({
        success: false,
        error: result.error
      });
    }

    // Opcionalmente subir a storage
    let storageUrl = null;
    if (req.body.saveToStorage) {
      const uploadResult = await excelService.uploadToStorage(result, data);
      if (uploadResult.success) {
        storageUrl = uploadResult.url;
      }
    }

    res.status(200).json({
      success: true,
      data: {
        fileName: result.fileName,
        formato: result.formato,
        base64Data: result.fileData,
        storageUrl,
        dataHash: result.dataHash
      },
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('[API] ❌ Error:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error interno del servidor'
    });
  }
});

/**
 * Listar plantillas disponibles
 */
app.get('/api/plantillas', async (req, res) => {
  try {
    const result = await storageUtils.listTemplates();

    if (!result.success) {
      return res.status(500).json({
        success: false,
        error: result.error
      });
    }

    const plantillas = result.templates.map(template => ({
      nombre: template.name,
      tamaño: template.metadata?.size || 0,
      fechaModificacion: template.updated_at,
      tipo: template.name.endsWith('.xlsx') ? 'excel' : template.name.endsWith('.csv') ? 'mapping' : 'unknown'
    }));

    res.status(200).json({
      success: true,
      plantillas,
      total: plantillas.length,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('[API] ❌ Error listando plantillas:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error obteniendo plantillas'
    });
  }
});

/**
 * Obtener historial de documentos generados
 */
app.get('/api/documentos/:pacienteId?', async (req, res) => {
  try {
    const { pacienteId } = req.params;
    const { formato, limite = 50 } = req.query;

    let query = req.app.locals.supabase
      .from('documentos_generados_sumate')
      .select('*')
      .order('fecha_generacion', { ascending: false })
      .limit(parseInt(limite));

    if (pacienteId) {
      query = query.eq('paciente_id', pacienteId);
    }

    if (formato) {
      query = query.eq('formato', formato);
    }

    const { data, error } = await query;

    if (error) {
      throw error;
    }

    res.status(200).json({
      success: true,
      documentos: data,
      total: data.length,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('[API] ❌ Error obteniendo historial:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error obteniendo historial de documentos'
    });
  }
});

/**
 * Descargar documento por ID
 */
app.get('/api/descargar/:documentId', async (req, res) => {
  try {
    const { documentId } = req.params;

    // Obtener metadata del documento
    const { data: documento, error } = await req.app.locals.supabase
      .from('documentos_generados_sumate')
      .select('*')
      .eq('id', documentId)
      .single();

    if (error || !documento) {
      return res.status(404).json({
        success: false,
        error: 'Documento no encontrado'
      });
    }

    // Incrementar contador de descargas
    await documentUtils.incrementDownloadCount(documentId);

    // Obtener URL del archivo
    const urlResult = await storageUtils.getPublicUrl(documento.storage_path);

    if (!urlResult.success) {
      return res.status(500).json({
        success: false,
        error: 'Error obteniendo URL de descarga'
      });
    }

    res.redirect(urlResult.url);

  } catch (error) {
    console.error('[API] ❌ Error en descarga:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error procesando descarga'
    });
  }
});

// =============================================
// ENDPOINTS DE ESTADÍSTICAS
// =============================================

/**
 * Estadísticas del servicio
 */
app.get('/api/stats', async (req, res) => {
  try {
    // Obtener stats de documentos
    const { data: stats, error } = await req.app.locals.supabase
      .from('documentos_generados_sumate')
      .select('formato, fecha_generacion')
      .gte('fecha_generacion', new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString());

    if (error) {
      throw error;
    }

    const now = new Date();
    const oneDayAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    const oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

    const totalDocuments = stats.length;
    const documentsToday = stats.filter(doc => new Date(doc.fecha_generacion) > oneDayAgo).length;
    const documentsThisWeek = stats.filter(doc => new Date(doc.fecha_generacion) > oneWeekAgo).length;

    // Agrupar por formato
    const porFormato = stats.reduce((acc, doc) => {
      acc[doc.formato] = (acc[doc.formato] || 0) + 1;
      return acc;
    }, {});

    res.status(200).json({
      success: true,
      stats: {
        totales: {
          documentos: totalDocuments,
          hoy: documentsToday,
          semana: documentsThisWeek
        },
        porFormato,
        sistema: {
          uptime: process.uptime(),
          memoria: {
            usado: Math.round(process.memoryUsage().heapUsed / 1024 / 1024),
            total: Math.round(process.memoryUsage().heapTotal / 1024 / 1024)
          }
        }
      },
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('[STATS] ❌ Error obteniendo estadísticas:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error obteniendo estadísticas'
    });
  }
});

// =============================================
// MANEJO DE ERRORES
// =============================================

app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: 'Endpoint no encontrado',
    path: req.path,
    method: req.method
  });
});

app.use((err, req, res, next) => {
  console.error('[ERROR]', err.stack);
  res.status(500).json({
    success: false,
    error: 'Error interno del servidor'
  });
});

// =============================================
// INICIAR SERVIDOR
// =============================================

async function startServer() {
  try {
    // Verificar conexiones
    const supabaseCheck = await checkSupabaseConnection();

    if (!supabaseCheck.success) {
      console.error('[INIT] ❌ Error conectando a Supabase:', supabaseCheck.error);
    } else {
      console.log('[INIT] ✅ Supabase conectado');
    }

    // Verificar storage
    const templatesCheck = await storageUtils.listTemplates();
    if (templatesCheck.success) {
      console.log(`[INIT] ✅ Storage verificado - ${templatesCheck.templates.length} plantillas disponibles`);
    } else {
      console.warn(`[INIT] ⚠️ Warning con storage: ${templatesCheck.error}`);
    }

    // Adjuntar Supabase a la aplicación para uso en rutas
    app.locals.supabase = require('./src/config/supabase');

    const server = app.listen(PORT, () => {
      console.log(`🚀 Constructor de Documentos Sumate ejecutándose en puerto ${PORT}`);
      console.log(`📊 Health check: http://localhost:${PORT}/health`);
      console.log(`📈 Estadísticas: http://localhost:${PORT}/api/stats`);
      console.log(`🔗 Endpoints disponibles:`);
      console.log(`   • POST /webhook/generar-documento - Webhook principal`);
      console.log(`   • POST /api/generar-documento - API directa`);
      console.log(`   • GET  /api/plantillas - Listar plantillas`);
      console.log(`   • GET  /api/documentos - Historial de documentos`);
      console.log('========================================');
    });

    // Manejo de cierre graceful
    process.on('SIGTERM', () => {
      console.log('[SHUTDOWN] Recibida señal SIGTERM, cerrando servidor...');
      server.close(() => {
        console.log('[SHUTDOWN] Servidor cerrado correctamente');
        process.exit(0);
      });
    });

  } catch (error) {
    console.error('[INIT] ❌ Error inicializando servidor:', error.message);
    process.exit(1);
  }
}

startServer();