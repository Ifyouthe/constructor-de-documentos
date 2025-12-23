// =============================================
// SERVICIO DE GENERACIÓN DE WORD - SUMATE
// Basado en la implementación de Nexus clinicalNoteService
// =============================================

const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const crypto = require('crypto');
const { storageUtils, documentUtils } = require('../../config/supabase');

class WordService {
  constructor() {
    console.log('[WORD-SERVICE] ✅ Servicio inicializado');
  }

  /**
   * Generar documento Word usando plantillas de Supabase Storage
   */
  async generateWord(data, template) {
    try {
      // Filtrar datos no nulos
      const filteredData = this.filterNonNullData(data);

      if (Object.keys(filteredData).length === 0) {
        return { success: false, error: 'No hay datos válidos para procesar' };
      }

      console.log(`[WORD-SERVICE] 📝 Generando documento Word: ${template}`);

      // Auto-detectar plantilla Word
      let templateName = template;
      if (!templateName.endsWith('.doc') && !templateName.endsWith('.docx')) {
        templateName = `${template}.docx`;
      }

      // Si no se especifica template, usar la primera plantilla Word disponible
      if (!template || template === 'general') {
        const templatesResult = await storageUtils.listTemplates();
        if (templatesResult.success && templatesResult.templates.length > 0) {
          const wordFiles = templatesResult.templates.filter(t =>
            t.name.endsWith('.doc') || t.name.endsWith('.docx')
          );
          if (wordFiles.length > 0) {
            templateName = wordFiles[0].name;
          }
        }
      }

      const docBuffer = await this.createDocFromTemplate(filteredData, templateName);

      // Construir nombre de archivo
      const fileName = this.buildFileName(filteredData, templateName);
      const base64Data = docBuffer.toString('base64');

      // Crear hash de los datos
      const dataHash = this.createDataHash(filteredData);

      return {
        success: true,
        fileName,
        fileData: base64Data,
        buffer: docBuffer,
        template: templateName,
        dataHash
      };
    } catch (error) {
      console.error('[WORD-SERVICE] ❌ Error en generateWord:', error);
      return { success: false, error: `Error al generar el archivo Word: ${error.message}` };
    }
  }

  /**
   * Crear documento Word desde plantilla usando docxtemplater (como en Nexus)
   */
  async createDocFromTemplate(data, templateName) {
    try {
      console.log(`[WORD-SERVICE] 📥 Descargando plantilla: ${templateName}`);

      // Descargar plantilla desde Supabase Storage
      const templateResult = await storageUtils.downloadTemplate(templateName);

      if (!templateResult.success) {
        throw new Error(`Error descargando plantilla ${templateName}: ${templateResult.error}`);
      }

      // Convertir Blob a ArrayBuffer y luego a Buffer
      const arrayBuffer = await templateResult.data.arrayBuffer();
      const templateBuffer = Buffer.from(arrayBuffer);

      console.log(`[WORD-SERVICE] ✅ Plantilla cargada: ${templateName}`);

      // Cargar plantilla con PizZip
      const zip = new PizZip(templateBuffer);

      // Crear instancia de Docxtemplater
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // Preparar datos para el template (aplanado)
      const templateData = this.prepareTemplateData(data);

      console.log(`[WORD-SERVICE] 📊 Datos para template:`, Object.keys(templateData));

      // Setear los datos en el template
      doc.setData(templateData);

      try {
        // Renderizar el documento
        doc.render();
        console.log(`[WORD-SERVICE] ✅ Documento renderizado exitosamente`);
      } catch (error) {
        console.error('[WORD-SERVICE] ❌ Error renderizando template:', error);
        throw new Error(`Error renderizando template: ${error.message}`);
      }

      // Obtener el buffer del documento generado
      const docBuffer = doc.getZip().generate({
        type: 'nodebuffer',
        compression: 'DEFLATE',
      });

      return docBuffer;
    } catch (error) {
      console.error('[WORD-SERVICE] ❌ Error creando documento Word:', error);
      throw error;
    }
  }

  /**
   * Preparar datos para el template (como en Nexus)
   */
  prepareTemplateData(data) {
    // Crear objeto plano con todos los campos
    const flatData = this.flattenObject(data);
    const templateData = {};

    // Copiar todos los campos
    Object.keys(flatData).forEach(key => {
      templateData[key] = flatData[key] || '';
    });

    // Agregar campos adicionales comunes
    templateData.fecha = templateData.fecha || new Date().toLocaleDateString();
    templateData.nombre_completo = `${templateData.nombre || ''} ${templateData.apellido_paterno || templateData.apellido || ''}`.trim();

    console.log(`[WORD-SERVICE] 🔄 Template data preparado con ${Object.keys(templateData).length} campos`);

    return templateData;
  }

  /**
   * Procesar webhook y generar documento Word
   */
  async processWebhookData(data) {
    try {
      // Extraer template del JSON
      const template = data.template || 'general';
      const { template: _, ...dataSinTemplate } = data;

      console.log(`[WORD-SERVICE] 📨 Procesando webhook para template: ${template}`);

      const wordResult = await this.generateWord(dataSinTemplate, template);

      if (!wordResult.success) {
        return wordResult;
      }

      // Subir a Supabase Storage
      const uploadResult = await this.uploadToStorage(wordResult, dataSinTemplate);

      if (!uploadResult.success) {
        console.error('[WORD-SERVICE] ❌ Error subiendo a storage:', uploadResult.error);
      }

      // Enviar a N8N
      await this.sendToN8n(wordResult.fileData, wordResult.fileName, {
        template: wordResult.template,
        dataHash: wordResult.dataHash,
        storageUrl: uploadResult.url || null
      });

      return {
        success: true,
        template: wordResult.template,
        fileName: wordResult.fileName,
        storageUrl: uploadResult.url || null,
        dataHash: wordResult.dataHash
      };
    } catch (error) {
      console.error('[WORD-SERVICE] ❌ Error procesando webhook:', error);
      return { success: false, error: 'Error al procesar webhook' };
    }
  }

  /**
   * Subir documento generado a Supabase Storage
   */
  async uploadToStorage(wordResult, originalData) {
    try {
      const fileName = wordResult.fileName;
      const buffer = wordResult.buffer;

      // Subir a storage
      const uploadResult = await storageUtils.uploadGeneratedDocument(fileName, buffer, {
        template: wordResult.template,
        dataHash: wordResult.dataHash,
        originalData: JSON.stringify(originalData)
      });

      if (!uploadResult.success) {
        return uploadResult;
      }

      // Obtener URL pública
      const urlResult = await storageUtils.getPublicUrl(fileName);

      // Guardar metadata en base de datos
      const metadataResult = await documentUtils.saveDocumentMetadata({
        pacienteId: originalData.paciente_id || originalData.id || null,
        formato: 'word',
        numeroExpediente: originalData.numero_de_expediente || originalData.expediente || null,
        waId: originalData.wa_id || null,
        storagePath: fileName,
        nombreArchivo: fileName,
        dataHash: wordResult.dataHash
      });

      if (!metadataResult.success) {
        console.warn('[WORD-SERVICE] ⚠️ Error guardando metadata:', metadataResult.error);
      }

      return {
        success: true,
        url: urlResult.url,
        path: fileName,
        metadata: metadataResult.data
      };
    } catch (error) {
      console.error('[WORD-SERVICE] ❌ Error en uploadToStorage:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Enviar archivo a N8N
   */
  async sendToN8n(base64Data, fileName, metadata = {}) {
    try {
      const webhookUrl = process.env.N8N_WEBHOOK_URL;

      if (!webhookUrl) {
        throw new Error('N8N_WEBHOOK_URL no configurada');
      }

      console.log(`[WORD-SERVICE] 📤 Enviando a N8N: ${fileName}`);

      const payload = {
        fileName: fileName,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        base64: base64Data,
        metadata: {
          generatedAt: new Date().toISOString(),
          source: 'constructor-documentos-sumate',
          type: 'word',
          ...metadata
        }
      };

      const response = await require('axios').post(webhookUrl, payload, {
        timeout: 30000,
        headers: {
          'Content-Type': 'application/json',
          'User-Agent': 'Constructor-Documentos-Sumate/1.0'
        }
      });

      console.log(`[WORD-SERVICE] ✅ Enviado exitosamente a N8N`);
    } catch (error) {
      console.error(`[WORD-SERVICE] ❌ Error enviando a N8N:`, error.message);
      throw error;
    }
  }

  // ===== MÉTODOS UTILITARIOS =====

  /**
   * Aplanar objeto a notación por puntos
   */
  flattenObject(obj, parentKey = '', result = {}) {
    if (!obj || typeof obj !== 'object') return result;

    for (const [key, value] of Object.entries(obj)) {
      const newKey = parentKey ? `${parentKey}.${key}` : key;
      const isPlainObject = value && typeof value === 'object' &&
        !Array.isArray(value) && !(value instanceof Date);

      if (isPlainObject) {
        this.flattenObject(value, newKey, result);
      } else {
        result[newKey] = value;
      }
    }

    return result;
  }

  /**
   * Construir nombre de archivo
   */
  buildFileName(data, templateName) {
    const today = new Date();
    const formattedDate = this.formatDateForFilenameDDMMYYYY(today);

    const baseName = templateName.replace(/\.(doc|docx)$/i, '');
    const nombre = data.nombre || 'SIN_NOMBRE';
    const apellido = data.apellido_paterno || data.apellido || '';
    const codigo = data.codigo || data.id || 'SIN_CODIGO';

    const nombreCompleto = `${nombre} ${apellido}`.trim();
    const nombreSanitized = this.sanitizeForFilenameUpper(nombreCompleto);
    const codigoSanitized = this.sanitizeForFilenameUpper(codigo);

    const extension = templateName.toLowerCase().endsWith('.docx') ? '.docx' : '.doc';

    return `SUMATE_${baseName}_${nombreSanitized}_${codigoSanitized}_${formattedDate}${extension}`;
  }

  /**
   * Filtrar datos no nulos
   */
  filterNonNullData(data) {
    const filtered = {};
    for (const [key, value] of Object.entries(data)) {
      if (value !== null && value !== undefined && value !== '') {
        filtered[key] = value;
      }
    }
    return filtered;
  }

  /**
   * Crear hash de datos
   */
  createDataHash(data) {
    const hash = crypto.createHash('md5');
    hash.update(JSON.stringify(data, Object.keys(data).sort()));
    return hash.digest('hex');
  }

  // Métodos de formato
  formatDateForFilenameDDMMYYYY(date) {
    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const yyyy = String(date.getFullYear());
    return `${dd}-${mm}-${yyyy}`;
  }

  sanitizeForFilenameUpper(value) {
    const base = String(value ?? '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]+/g, '_')
      .replace(/_{2,}/g, '_')
      .replace(/^_+|_+$/g, '');
    return base.toUpperCase() || 'SIN_VALOR';
  }
}

module.exports = new WordService();