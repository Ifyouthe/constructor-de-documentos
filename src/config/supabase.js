// =============================================
// CONFIGURACIÓN DE SUPABASE - CONSTRUCTOR DOCUMENTOS SUMATE
// =============================================

const { createClient } = require('@supabase/supabase-js');

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

// Log de configuración
console.log('🔌 Constructor de Documentos Sumate conectado a Supabase');
console.log(`📊 Supabase URL: ${SUPABASE_URL}`);
console.log(`🔑 Usando ${SUPABASE_ANON_KEY ? 'ANON KEY' : 'SERVICE KEY'} para operaciones backend`);

/**
 * Verificar conexión a Supabase
 */
async function checkSupabaseConnection() {
  try {
    console.log('[SUPABASE] 🔍 Probando conexión con query simple...');

    // Probar con una query más simple primero
    const { data, error } = await supabaseService
      .from('documentos_generados_sumate')
      .select('*')
      .limit(1);

    console.log('[SUPABASE] 🔍 Respuesta de query:', { data, error });

    if (error && error.code !== 'PGRST116') {
      throw error;
    }

    console.log('[SUPABASE] ✅ Conexión verificada exitosamente');
    return { success: true };
  } catch (error) {
    console.error('[SUPABASE] ❌ Error verificando conexión:', error.message);
    console.error('[SUPABASE] 🔍 Detalles del error:', error);
    console.error('[SUPABASE] 🔍 URL:', SUPABASE_URL);
    console.error('[SUPABASE] 🔍 Key type:', SUPABASE_ANON_KEY ? 'ANON' : 'SERVICE');
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
      console.error('[STORAGE] Error descargando plantilla:', error.message || error);
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
      console.error('[STORAGE] Error subiendo documento:', error.message);
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
      console.error('[STORAGE] Error obteniendo URL pública:', error.message);
      return { success: false, error: error.message };
    }
  },

  /**
   * Listar plantillas disponibles
   */
  async listTemplates() {
    try {
      const bucketName = process.env.SUPABASE_BUCKET_TEMPLATES || 'plantillas-documentos';
      console.log(`[STORAGE] 📂 Listando plantillas del bucket: ${bucketName}`);
      console.log(`[STORAGE] 🔑 Variables de entorno: TEMPLATES=${process.env.SUPABASE_BUCKET_TEMPLATES}, GENERATED=${process.env.SUPABASE_BUCKET_GENERATED}`);

      const { data, error } = await supabaseService.storage
        .from(bucketName)
        .list();

      if (error) {
        console.error('[STORAGE] ❌ Error listando plantillas:', error);
        throw error;
      }

      console.log(`[STORAGE] 📋 Encontrados ${data.length} archivos en bucket ${bucketName}`);
      if (data.length > 0) {
        console.log('[STORAGE] 📁 Archivos encontrados:', data.map(f => f.name));
      }

      return { success: true, templates: data };
    } catch (error) {
      console.error('[STORAGE] Error listando plantillas:', error.message);
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
      console.error('[DATABASE] Error guardando metadata:', error.message);
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
      console.error('[DATABASE] Error buscando documento:', error.message);
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
      console.error('[DATABASE] Error actualizando descargas:', error.message);
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