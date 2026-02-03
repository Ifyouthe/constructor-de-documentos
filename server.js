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

// Importar servicios (ya vienen como instancias)
const excelService = require('./src/services/excel/excelService');
const wordService = require('./src/services/word/wordService');
const multipleDocumentsService = require('./src/services/multipleDocumentsService');

// Sistema de logging
const { createLogger, summarizeRequest } = require('./src/utils/logger');
const log = createLogger('SERVER');

const app = express();
const PORT = process.env.PORT || 3003;

// Configurar trust proxy para Traefik
app.set('trust proxy', true);

log.info('Iniciando Constructor de Documentos SUMATE');
log.info(`Puerto: ${PORT} | Entorno: ${process.env.NODE_ENV || 'development'}`);

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

// Extraer datos_prospecto para endpoints de multiples documentos
const normalizeDatosProspecto = (body) => {
  if (body && typeof body.datos_prospecto === 'object' && body.datos_prospecto !== null) {
    return body.datos_prospecto;
  }

  return null;
};

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
    log.error('Error en health check:', error.message);
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
    log.info('Request recibido: /webhook/generar-documento', summarizeRequest(req.body));

    // Detectar tipo de documento basado en formato o tipo especificado
    let formato = req.body.formato || req.body.template || 'general';

    // Si el template viene con el nombre completo del archivo, extraer el formato
    if (formato.includes('Fichadeidentificaciondelobligadosolidario')) {
      formato = 'obligado_solidario';
    } else if (formato.includes('Visita domiciliaria')) {
      formato = 'visita_domiciliaria';
    } else if (formato.includes('aval')) {
      formato = 'ficha_aval';
    }

    const documentType = req.body.type ||
      (formato === 'obligado_solidario' || formato === 'obligado' || formato === 'ficha_obligado' ||
       formato === 'visita_domiciliaria' || formato === 'ficha_aval' ? 'word' : 'excel');

    let result;

    if (documentType === 'word') {
      // Configurar template correcto para Word
      const dataConTemplate = { ...req.body };
      if (formato === 'obligado_solidario' || formato === 'obligado' || formato === 'ficha_obligado') {
        dataConTemplate.template = 'Fichadeidentificaciondelobligadosolidarioconetiquetas.docx';
      } else if (formato === 'visita_domiciliaria') {
        dataConTemplate.template = 'Visita domiciliaria con etiquetas.docx';
      } else if (formato === 'ficha_aval') {
        dataConTemplate.template = 'ficha_de_identificacion_del_aval_con_etiquetas.docx';
      }

      result = await wordService.processWebhookData(dataConTemplate);
    } else {
      result = await excelService.processWebhookData(req.body);
    }

    if (!result.success) {
      log.error('Error procesando documento:', result.error);
      return res.status(400).json({
        success: false,
        error: result.error,
        timestamp: new Date().toISOString()
      });
    }

    log.info('Documento generado exitosamente:', result.fileName);

    // Verificar que el buffer existe
    if (!result.buffer) {
      log.error('No se genero buffer del documento');
      return res.status(500).json({
        success: false,
        error: 'No se pudo generar el archivo',
        timestamp: new Date().toISOString()
      });
    }

    // Devolver siempre el archivo construido directamente
    const mimeType = documentType === 'word'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    // Sanitizar nombre de archivo para evitar problemas con caracteres especiales
    const safeFileName = result.fileName.replace(/[^a-zA-Z0-9._-]/g, '_');

    res.setHeader('Content-Type', mimeType);
    res.setHeader('Content-Disposition', `attachment; filename="${safeFileName}"`);
    res.setHeader('Content-Length', result.buffer.length);
    res.send(result.buffer);

  } catch (error) {
    log.error('Error interno en webhook:', error.message);
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
    log.info('Request recibido: /api/generar-documento', summarizeRequest(req.body));

    // IMPORTANTE: Detectar formato desde 'template' (lo que envía N8N)
    let formato = req.body.template || req.body.formato || 'general';
    let tipoDocumento = req.body.type; // 'excel', 'word', etc

    // Extraer data - quitar campos de control
    let data;
    if (req.body.data) {
      // Si viene con wrapper 'data'
      data = req.body.data;
    } else {
      // Si viene todo junto (como lo envía N8N), quitar campos de control
      const { type, template, formato: f, saveToStorage, ...restData } = req.body;
      data = restData;
    }

    if (!data || Object.keys(data).length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Se requieren datos para generar el documento'
      });
    }

    log.info(`Generando documento - Template: ${formato}, Type: ${tipoDocumento}, Campos: ${Object.keys(data).length}`);

    let result;

    // Determinar si usar Word o Excel basado en el TEMPLATE (no en type)
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
          templateName = 'ficha_de_identificacion_del_aval_con_etiquetas.docx';
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
      // Usar el servicio correcto para upload
      const uploadService = (formato === 'obligado_solidario' || formato === 'obligado' ||
                            formato === 'ficha_obligado' || formato === 'visita_domiciliaria' ||
                            formato === 'ficha_aval') ? wordService : excelService;

      const uploadResult = await uploadService.uploadToStorage(result, data);
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
    log.error('Error en API generar-documento:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error interno del servidor'
    });
  }
});

/**
 * Generar múltiples documentos en una sola request
 */
app.post('/api/generar-multiples-documentos', async (req, res) => {
  try {
    log.info('Request recibido: /api/generar-multiples-documentos', {
      fichas: req.body.fichas_a_generar,
      tiene_datos: !!req.body.datos_prospecto
    });

    // Validar estructura de entrada
    const { fichas_a_generar } = req.body;
    const datos_prospecto = normalizeDatosProspecto(req.body);

    if (!fichas_a_generar || !Array.isArray(fichas_a_generar) || fichas_a_generar.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere un array "fichas_a_generar" con al menos una ficha'
      });
    }

    if (!datos_prospecto || typeof datos_prospecto !== 'object' || Object.keys(datos_prospecto).length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere un objeto "datos_prospecto" con los datos del cliente'
      });
    }

    // Validar fichas soportadas
    const fichasSoportadas = [
      'identificacion_cliente',
      'visita_domiciliaria',
      'evaluacion_economica_simple',
      'obligado_solidario',
      'aval',
      'seguimiento_previo',
      'scoring_con_hc',
      'scoring_sin_hc',
      'scoring',
      'seguimiento_credito'
    ];

    const fichasNoSoportadas = fichas_a_generar.filter(ficha => !fichasSoportadas.includes(ficha));

    if (fichasNoSoportadas.length > 0) {
      return res.status(400).json({
        success: false,
        error: `Fichas no soportadas: ${fichasNoSoportadas.join(', ')}`,
        fichas_soportadas: fichasSoportadas
      });
    }

    log.info(`Procesando ${fichas_a_generar.length} fichas para prospecto ${datos_prospecto.codigo_de_prospecto || datos_prospecto.id_expediente || 'SIN_CODIGO'}`);

    // Generar múltiples documentos
    const resultado = await multipleDocumentsService.generateMultipleDocuments(fichas_a_generar, datos_prospecto);

    if (!resultado.success) {
      log.error('Error en generacion multiple:', resultado.error);
      return res.status(400).json({
        success: false,
        error: resultado.error || 'Error al generar los documentos',
        detalles: resultado.detalles,
        errores: resultado.errores || [],
        metadata: resultado.metadata
      });
    }

    log.info(`Generacion completada: ${resultado.metadata.total_generados}/${resultado.metadata.total_solicitados} documentos`);

    // Respuesta con todos los documentos generados
    res.status(200).json({
      success: true,
      data: {
        documentos_generados: resultado.documentos_generados,
        errores: resultado.errores,
        resumen: {
          total_solicitados: resultado.metadata.total_solicitados,
          total_generados: resultado.metadata.total_generados,
          total_errores: resultado.metadata.total_errores,
          porcentaje_exito: Math.round((resultado.metadata.total_generados / resultado.metadata.total_solicitados) * 100)
        }
      },
      timestamp: resultado.metadata.timestamp
    });

  } catch (error) {
    log.error('Error critico en API multiple:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error interno del servidor',
      detalles: error.message
    });
  }
});

/**
 * Generar múltiples documentos y descargar como ZIP
 */
app.post('/webhook/generar-multiples-documentos-zip', async (req, res) => {
  try {
    log.info('Request recibido: /webhook/generar-multiples-documentos-zip', summarizeRequest(req.body));

    const { fichas_a_generar } = req.body;
    const datos_prospecto = normalizeDatosProspecto(req.body);

    // Validaciones básicas
    if (!fichas_a_generar || !Array.isArray(fichas_a_generar)) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere array "fichas_a_generar"',
        timestamp: new Date().toISOString()
      });
    }

    if (!datos_prospecto || typeof datos_prospecto !== 'object' || Object.keys(datos_prospecto).length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere objeto "datos_prospecto"',
        timestamp: new Date().toISOString()
      });
    }

    // Generar documentos
    const resultado = await multipleDocumentsService.generateMultipleDocuments(fichas_a_generar, datos_prospecto);

    if (!resultado.success) {
      log.error('Error en webhook-multiple-zip:', resultado.error);
      return res.status(400).json({
        success: false,
        error: resultado.error,
        errores: resultado.errores,
        timestamp: new Date().toISOString()
      });
    }

    // Si solo hay un documento, descargarlo directamente
    if (resultado.documentos_generados.length === 1) {
      const doc = resultado.documentos_generados[0];
      const buffer = Buffer.from(doc.fileData, 'base64');

      const mimeType = doc.tipo_ficha.includes('visita_domiciliaria') ||
                      doc.tipo_ficha.includes('obligado_solidario') ||
                      doc.tipo_ficha.includes('aval')
                      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

      res.setHeader('Content-Type', mimeType);
      res.setHeader('Content-Disposition', `attachment; filename="${doc.fileName}"`);
      res.setHeader('Content-Length', buffer.length);
      return res.send(buffer);
    }

    // Para múltiples documentos, crear un ZIP
    const AdmZip = require('adm-zip');
    const zip = new AdmZip();

    for (const doc of resultado.documentos_generados) {
      const buffer = Buffer.from(doc.fileData, 'base64');
      zip.addFile(doc.fileName, buffer);
    }

    const zipBuffer = zip.toBuffer();
    const zipFileName = `${datos_prospecto.primer_apellido || 'DOCUMENTOS'}_${datos_prospecto.codigo_de_prospecto || datos_prospecto.id_expediente || 'MULTIPLE'}.zip`;

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${zipFileName}"`);
    res.setHeader('Content-Length', zipBuffer.length);
    res.send(zipBuffer);

    log.info(`ZIP generado con ${resultado.documentos_generados.length} documentos`);

  } catch (error) {
    log.error('Error interno en webhook-multiple-zip:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error interno del servidor',
      timestamp: new Date().toISOString()
    });
  }
});

/**
 * Webhook para generar múltiples documentos (compatible con N8N)
 */
app.post('/webhook/generar-multiples-documentos', async (req, res) => {
  try {
    log.info('Request recibido: /webhook/generar-multiples-documentos', summarizeRequest(req.body));

    const { fichas_a_generar } = req.body;
    const datos_prospecto = normalizeDatosProspecto(req.body);

    // Validaciones básicas
    if (!fichas_a_generar || !Array.isArray(fichas_a_generar)) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere array "fichas_a_generar"',
        timestamp: new Date().toISOString()
      });
    }

    if (!datos_prospecto || typeof datos_prospecto !== 'object' || Object.keys(datos_prospecto).length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Se requiere objeto "datos_prospecto"',
        timestamp: new Date().toISOString()
      });
    }

    // Generar documentos
    const resultado = await multipleDocumentsService.generateMultipleDocuments(fichas_a_generar, datos_prospecto);

    if (!resultado.success) {
      log.error('Error en webhook-multiple:', resultado.error);
      return res.status(400).json({
        success: false,
        error: resultado.error,
        errores: resultado.errores,
        timestamp: new Date().toISOString()
      });
    }

    // Para webhook, devolver estructura compatible con N8N
    // Nota: Los archivos se devuelven como base64, N8N puede procesarlos individualmente
    const respuesta = {
      success: true,
      documentos: resultado.documentos_generados.map(doc => ({
        tipo_ficha: doc.tipo_ficha,
        fileName: doc.fileName,
        mimeType: doc.tipo_ficha.includes('visita_domiciliaria') ||
                  doc.tipo_ficha.includes('obligado_solidario') ||
                  doc.tipo_ficha.includes('aval')
                  ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                  : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        base64Data: doc.fileData,
        metadata: doc.metadata
      })),
      errores: resultado.errores,
      resumen: {
        total_solicitados: resultado.metadata.total_solicitados,
        total_generados: resultado.metadata.total_generados,
        total_errores: resultado.metadata.total_errores
      },
      timestamp: resultado.metadata.timestamp
    };

    log.info(`${resultado.metadata.total_generados} documentos generados exitosamente`);

    res.status(200).json(respuesta);

  } catch (error) {
    log.error('Error interno en webhook-multiple:', error.message);
    res.status(500).json({
      success: false,
      error: 'Error interno del servidor',
      timestamp: new Date().toISOString()
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
    log.error('Error listando plantillas:', error.message);
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
    log.error('Error obteniendo historial:', error.message);
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
    log.error('Error en descarga:', error.message);
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
    log.error('Error obteniendo estadisticas:', error.message);
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
  log.error('Error no manejado:', err.message);
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
      log.error('Error conectando a Supabase:', supabaseCheck.error);
    } else {
      log.info('Supabase conectado');
    }

    // Verificar storage
    const templatesCheck = await storageUtils.listTemplates();
    if (templatesCheck.success) {
      log.info(`Storage verificado - ${templatesCheck.templates.length} plantillas disponibles`);
    } else {
      log.warn('Warning con storage:', templatesCheck.error);
    }

    // Adjuntar Supabase a la aplicación para uso en rutas
    app.locals.supabase = require('./src/config/supabase');

    const server = app.listen(PORT, () => {
      log.info(`Servidor ejecutandose en puerto ${PORT}`);
      log.info(`Health: http://localhost:${PORT}/health | Stats: http://localhost:${PORT}/api/stats`);
    });

    // Manejo de cierre graceful
    process.on('SIGTERM', () => {
      log.info('Recibida senal SIGTERM, cerrando servidor...');
      server.close(() => {
        log.info('Servidor cerrado correctamente');
        process.exit(0);
      });
    });

  } catch (error) {
    log.error('Error inicializando servidor:', error.message);
    process.exit(1);
  }
}

startServer();
