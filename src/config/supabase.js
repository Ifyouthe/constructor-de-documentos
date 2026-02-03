// =============================================
// CONFIGURACIÓN DE SUPABASE - CONSTRUCTOR DOCUMENTOS SUMATE
// =============================================

const { createClient } = require('@supabase/supabase-js');

// Logger simple para supabase (evitar dependencia circular)
const log = {
  info: (msg, data) => console.log(`[INFO] [SUPABASE] ${msg}`, data !== undefined ? data : ''),
  error: (msg, data) => console.error(`[ERROR] [SUPABASE] ${msg}`, data !== undefined ? data : ''),
  warn: (msg, data) => console.warn(`[WARN] [SUPABASE] ${msg}`, data !== undefined ? data : ''),
  debug: (msg, data) => process.env.LOG_LEVEL === 'DEBUG' && console.log(`[DEBUG] [SUPABASE] ${msg}`, data !== undefined ? data : '')
};

// Configuración desde variables de entorno
const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY;

// Validación crítica
if (!SUPABASE_URL) {
  throw new Error('SUPABASE_URL no está definida. Verifica tu archivo .env');
}
if (!SUPABASE_SERVICE_KEY && !SUPABASE_ANON_KEY) {
  throw new Error('Necesitas SUPABASE_SERVICE_ROLE_KEY o SUPABASE_ANON_KEY. Verifica tu archivo .env');
}

/**
 * Cliente Supabase con ANON KEY para operaciones
 * Usar ANON_KEY si SERVICE_KEY falla
 */
const supabaseService = createClient(SUPABASE_URL, SUPABASE_ANON_KEY || SUPABASE_SERVICE_KEY, {
  auth: {
    autoRefreshToken: false,
    persistSession: false
  }
});

/**
 * Cliente Supabase con ANON KEY para operaciones limitadas
 * (respeta RLS - Row Level Security)
 */
const supabaseAnon = SUPABASE_ANON_KEY ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    autoRefreshToken: false,
    persistSession: false
  }
}) : null;

log.info(`Conectado a Supabase | Key: ${SUPABASE_ANON_KEY ? 'ANON' : 'SERVICE'}`);

/**
 * Verificar conexión a Supabase
 */
async function checkSupabaseConnection() {
  try {
    const { data, error } = await supabaseService
      .from('documentos_generados_sumate')
      .select('*')
      .limit(1);

    if (error && error.code !== 'PGRST116') {
      throw error;
    }

    return { success: true };
  } catch (error) {
    log.error('Error verificando conexion:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Funciones para Storage de Supabase
 */
const storageUtils = {
  /**
   * Descargar plantilla desde storage
   */
  async downloadTemplate(templateName) {
    try {
      const { data, error } = await supabaseService.storage
        .from(process.env.SUPABASE_BUCKET_TEMPLATES || 'plantillas-documentos')
        .download(templateName);

      if (error) throw error;

      return { success: true, data };
    } catch (error) {
      log.error('Error descargando plantilla:', error.message || error);
      return { success: false, error: error.message || error.toString() };
    }
  },

  /**
   * Subir documento generado al storage
   */
  async uploadGeneratedDocument(fileName, fileBuffer, metadata = {}) {
    try {
      const { data, error } = await supabaseService.storage
        .from(process.env.SUPABASE_BUCKET_GENERATED || 'documentos-generados')
        .upload(fileName, fileBuffer, {
          cacheControl: '3600',
          upsert: true,
          metadata: {
            ...metadata,
            generatedAt: new Date().toISOString(),
            source: 'constructor-documentos-sumate'
          }
        });

      if (error) throw error;

      return { success: true, data };
    } catch (error) {
      log.error('Error subiendo documento:', error.message);
      return { success: false, error: error.message };
    }
  },

  /**
   * Obtener URL pública de un documento
   */
  async getPublicUrl(fileName, bucket = null) {
    try {
      const bucketName = bucket || process.env.SUPABASE_BUCKET_GENERATED || 'documentos-generados';

      const { data } = supabaseService.storage
        .from(bucketName)
        .getPublicUrl(fileName);

      return { success: true, url: data.publicUrl };
    } catch (error) {
      log.error('Error obteniendo URL publica:', error.message);
      return { success: false, error: error.message };
    }
  },

  /**
   * Listar plantillas disponibles
   */
  async listTemplates() {
    try {
      const bucketName = process.env.SUPABASE_BUCKET_TEMPLATES || 'plantillas-documentos';

      const { data, error } = await supabaseService.storage
        .from(bucketName)
        .list();

      if (error) {
        throw error;
      }

      return { success: true, templates: data };
    } catch (error) {
      log.error('Error listando plantillas:', error.message);
      return { success: false, error: error.message };
    }
  }
};

/**
 * Funciones para la tabla documentos_generados_sumate
 */
const documentUtils = {
  /**
   * Guardar metadata de documento generado
   */
  async saveDocumentMetadata(metadata) {
    try {
      const { data, error } = await supabaseService
        .from('documentos_generados_sumate')
        .insert({
          paciente_id: metadata.pacienteId,
          formato: metadata.formato,
          numero_de_expediente: metadata.numeroExpediente,
          wa_id: metadata.waId,
          storage_path: metadata.storagePath,
          nombre_archivo: metadata.nombreArchivo,
          data_hash: metadata.dataHash,
          fecha_generacion: new Date().toISOString()
        })
        .select()
        .single();

      if (error) throw error;

      return { success: true, data };
    } catch (error) {
      log.error('Error guardando metadata:', error.message);
      return { success: false, error: error.message };
    }
  },

  /**
   * Buscar documento existente
   */
  async findExistingDocument(pacienteId, formato) {
    try {
      const { data, error } = await supabaseService
        .from('documentos_generados_sumate')
        .select('*')
        .eq('paciente_id', pacienteId)
        .eq('formato', formato)
        .order('fecha_generacion', { ascending: false })
        .limit(1);

      if (error) throw error;

      return { success: true, document: data[0] || null };
    } catch (error) {
      log.error('Error buscando documento:', error.message);
      return { success: false, error: error.message };
    }
  },

  /**
   * Actualizar contador de descargas
   */
  async incrementDownloadCount(documentId) {
    try {
      const { data, error } = await supabaseService
        .from('documentos_generados_sumate')
        .update({
          numero_descargas: supabaseService.sql`numero_descargas + 1`,
          ultima_descarga: new Date().toISOString()
        })
        .eq('id', documentId)
        .select()
        .single();

      if (error) throw error;

      return { success: true, data };
    } catch (error) {
      log.error('Error actualizando descargas:', error.message);
      return { success: false, error: error.message };
    }
  }
};

// Exportar supabaseService como default para compatibilidad
module.exports = supabaseService;
module.exports.supabaseService = supabaseService;
module.exports.supabaseAnon = supabaseAnon;
module.exports.checkSupabaseConnection = checkSupabaseConnection;
module.exports.storageUtils = storageUtils;
module.exports.documentUtils = documentUtils;