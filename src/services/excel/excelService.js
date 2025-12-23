// =============================================
// SERVICIO DE GENERACIÓN DE EXCEL - SUMATE
// Basado en la implementación de Nexus con adaptaciones para Sumate
// =============================================

const ExcelJS = require('exceljs');
const axios = require('axios');
const crypto = require('crypto');
const { storageUtils, documentUtils } = require('../../config/supabase');
const mappingService = require('../../utils/mappingService');

class ExcelService {
  constructor() {
    console.log('[EXCEL-SERVICE] ✅ Servicio inicializado');
  }

  /**
   * Generar documento Excel usando plantillas de Supabase Storage
   */
  async generateExcel(data, formato = 'general') {
    try {
      // Si viene como array, tomar el primer elemento
      let dataToProcess = data;
      if (Array.isArray(data) && data.length > 0) {
        dataToProcess = data[0];
      }

      const filteredData = this.filterNonNullData(dataToProcess);

      if (Object.keys(filteredData).length === 0) {
        return { success: false, error: 'No hay datos válidos para procesar' };
      }

      // No validar formato ya que ahora es dinámico basado en templates disponibles

      console.log(`[EXCEL-SERVICE] 🔄 Generando documento formato: ${formato}`);

      const workbook = await this.createWorkbookFromTemplate(filteredData, formato);
      const buffer = await workbook.xlsx.writeBuffer();

      // Construir nombre de archivo
      const fileName = this.buildFileName(filteredData, formato);
      const base64Data = buffer.toString('base64');

      // Crear hash de los datos para detección de cambios
      const dataHash = this.createDataHash(filteredData);

      return {
        success: true,
        fileName,
        fileData: base64Data,
        buffer,
        formato,
        dataHash
      };
    } catch (error) {
      console.error('[EXCEL-SERVICE] ❌ Error en generateExcel:', error);
      return { success: false, error: `Error al generar el archivo Excel: ${error.message}` };
    }
  }

  /**
   * Procesar webhook y generar documento
   */
  async processWebhookData(data) {
    try {
      // Extraer formato y template del JSON
      const formato = data.formato || data.template || 'general';
      const { formato: _, template: __, ...dataSinFormato } = data;

      // Preservar template en los datos para que llegue a createWorkbookFromTemplate
      if (data.template) {
        dataSinFormato.template = data.template;
      }

      console.log(`[EXCEL-SERVICE] 📨 Procesando webhook para formato: ${formato}, template: ${data.template || 'auto'}`);

      const excelResult = await this.generateExcel(dataSinFormato, formato);

      if (!excelResult.success) {
        return excelResult;
      }

      // Subir a Supabase Storage
      const uploadResult = await this.uploadToStorage(excelResult, dataSinFormato);

      if (!uploadResult.success) {
        console.error('[EXCEL-SERVICE] ❌ Error subiendo a storage:', uploadResult.error);
      }

      // Enviar a N8N
      await this.sendToN8n(excelResult.fileData, excelResult.fileName, {
        formato: excelResult.formato,
        dataHash: excelResult.dataHash,
        storageUrl: uploadResult.url || null
      });

      return {
        success: true,
        formato: excelResult.formato,
        fileName: excelResult.fileName,
        storageUrl: uploadResult.url || null,
        dataHash: excelResult.dataHash
      };
    } catch (error) {
      console.error('[EXCEL-SERVICE] ❌ Error procesando webhook:', error);
      return { success: false, error: 'Error al procesar webhook' };
    }
  }

  /**
   * Subir documento generado a Supabase Storage
   */
  async uploadToStorage(excelResult, originalData) {
    try {
      const fileName = excelResult.fileName;
      const buffer = excelResult.buffer;

      // Subir a storage
      const uploadResult = await storageUtils.uploadGeneratedDocument(fileName, buffer, {
        formato: excelResult.formato,
        dataHash: excelResult.dataHash,
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
        formato: excelResult.formato,
        numeroExpediente: originalData.numero_de_expediente || originalData.expediente || null,
        waId: originalData.wa_id || null,
        storagePath: fileName,
        nombreArchivo: fileName,
        dataHash: excelResult.dataHash
      });

      if (!metadataResult.success) {
        console.warn('[EXCEL-SERVICE] ⚠️ Error guardando metadata:', metadataResult.error);
      }

      return {
        success: true,
        url: urlResult.url,
        path: fileName,
        metadata: metadataResult.data
      };
    } catch (error) {
      console.error('[EXCEL-SERVICE] ❌ Error en uploadToStorage:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Construir nombre de archivo según el formato
   */
  buildFileName(data, formato) {
    const formattedDate = this.formatDateForFilenameDDMMYYYY(new Date());
    const { nombreClienteUpper, codigoProspectoUpper } = this.buildNameParts(data);

    switch (formato) {
      case 'con_HC': {
        return `SUMATE_SCORING_CON_HC_${nombreClienteUpper}_${codigoProspectoUpper}_${formattedDate}.xlsx`;
      }
      case 'sin_HC': {
        return `SUMATE_SCORING_SIN_HC_${nombreClienteUpper}_${codigoProspectoUpper}_${formattedDate}.xlsx`;
      }
      case 'expediente_sumate': {
        const numeroExpediente = data?.numero_de_expediente || data?.expediente || 'SIN_EXPEDIENTE';
        const nombreCompleto = `${data?.nombre || ''} ${data?.apellido_paterno || data?.apellido || ''}`.trim() || 'SIN_NOMBRE';

        const expedienteUpper = this.sanitizeForFilenameUpper(numeroExpediente);
        const pacienteUpper = this.sanitizeForFilenameUpper(nombreCompleto);

        return `SUMATE_EXPEDIENTE_${expedienteUpper}_${pacienteUpper}_${formattedDate}.xlsx`;
      }
      case 'solicitud_credito': {
        return `SUMATE_SOLICITUD_CREDITO_${nombreClienteUpper}_${codigoProspectoUpper}_${formattedDate}.xlsx`;
      }
      case 'general':
      default: {
        return `SUMATE_DOCUMENTO_${nombreClienteUpper}_${codigoProspectoUpper}_${formattedDate}.xlsx`;
      }
    }
  }

  /**
   * Crear workbook desde plantilla almacenada en Supabase Storage
   */
  async createWorkbookFromTemplate(data, formato = 'general') {
    try {
      let templateName;
      let sheetName;

      // 1. Si se especifica un template directamente
      if (data.template) {
        templateName = data.template.endsWith('.xlsx') ? data.template : `${data.template}.xlsx`;

        // Mapear nombres de hojas según el template (como en Nexus)
        if (templateName.includes('SCORING_CON_HC')) {
          sheetName = 'Scoring del Cliente';
        } else if (templateName.includes('SCORING_SIN_HC')) {
          sheetName = 'Scoring del Cliente';
        } else if (templateName.includes('Formato_Editable_Listo')) {
          sheetName = 'Ficha de identificación';
        } else if (templateName.includes('seguimiento')) {
          sheetName = 'Hoja1'; // Ajustar según el nombre real
        } else {
          sheetName = 'Hoja1'; // Default
        }

        console.log(`[EXCEL-SERVICE] 📋 Usando template: ${templateName}, hoja: ${sheetName}`);
      }
      // 2. Si no hay template, mapear por formato
      else {
        switch (formato) {
          case 'con_HC':
            templateName = 'SCORING_CON_HC.xlsx';
            sheetName = 'Scoring del Cliente';
            break;
          case 'sin_HC':
            templateName = 'SCORING_SIN_HC.xlsx';
            sheetName = 'Scoring del Cliente';
            break;
          case 'general':
          default:
            templateName = 'Formato_Editable_Listo.xlsx';
            sheetName = 'Ficha de identificación';
            break;
        }
        console.log(`[EXCEL-SERVICE] 📋 Usando formato ${formato}: template=${templateName}, hoja=${sheetName}`);
      }

      console.log(`[EXCEL-SERVICE] 📥 Descargando plantilla: ${templateName}`);

      // Descargar plantilla desde Supabase Storage
      const templateResult = await storageUtils.downloadTemplate(templateName);

      if (!templateResult.success) {
        throw new Error(`Error descargando plantilla ${templateName}: ${templateResult.error}`);
      }

      // Convertir Blob a ArrayBuffer
      const arrayBuffer = await templateResult.data.arrayBuffer();

      // Cargar workbook desde buffer
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        throw new Error(`Hoja "${sheetName}" no encontrada en la plantilla`);
      }

      console.log(`[EXCEL-SERVICE] ✅ Plantilla cargada: ${templateName}`);

      // Cargar mappings y aplicar datos
      await mappingService.loadMappings(formato);
      const dataMapping = mappingService.createDataMapping(data, formato);
      const mappings = mappingService.getAllMappings(formato);

      console.log(`[EXCEL-SERVICE] 📊 DataMapping creado con ${dataMapping.size} entradas`);
      console.log(`[EXCEL-SERVICE] 📍 Mappings cargados: ${mappings.length}`);

      // Limpiar rellenos rojos de la plantilla
      this.clearAllRedFillsFromTemplate(worksheet);

      // Llenar celdas con datos
      this.fillCellsFromMappings(worksheet, dataMapping, mappings);

      // Aplicar reglas específicas por formato
      this.applyFormatSpecificRules(worksheet, formato, data);

      return workbook;
    } catch (error) {
      console.error('[EXCEL-SERVICE] ❌ Error creando workbook:', error);
      throw error;
    }
  }

  /**
   * Aplicar reglas específicas por formato
   */
  applyFormatSpecificRules(worksheet, formato, data) {
    switch (formato) {
      case 'sin_HC':
        // Vaciar únicamente el rango K7/K8 (merged)
        ['K7', 'K8'].forEach((addr) => {
          try {
            const cell = this.getMasterCell(worksheet, addr);
            if (cell) {
              cell.value = null;
            }
          } catch (_) {}
        });
        break;

      case 'expediente_sumate':
        try {
          const password = this.generateDynamicPassword();
          console.log(`[EXCEL-SERVICE] 🔒 Aplicando protección con contraseña al expediente Sumate`);

          worksheet.protect(password, {
            selectLockedCells: true,
            selectUnlockedCells: true,
          });

          console.log(`[EXCEL-SERVICE] ✅ Protección aplicada exitosamente`);
        } catch (protectionError) {
          console.error(`[EXCEL-SERVICE] ❌ Error aplicando protección:`, protectionError.message);
        }
        break;
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

      console.log(`[EXCEL-SERVICE] 📤 Enviando a N8N: ${fileName}`);

      const payload = {
        fileName: fileName,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        base64: base64Data,
        metadata: {
          generatedAt: new Date().toISOString(),
          source: 'constructor-documentos-sumate',
          ...metadata
        }
      };

      await axios.post(webhookUrl, payload, {
        timeout: 30000,
        headers: {
          'Content-Type': 'application/json',
          'User-Agent': 'Constructor-Documentos-Sumate/1.0'
        }
      });

      console.log(`[EXCEL-SERVICE] ✅ Enviado exitosamente a N8N`);
    } catch (error) {
      console.error(`[EXCEL-SERVICE] ❌ Error enviando a N8N:`, error.message);
      throw error;
    }
  }

  /**
   * Crear hash de datos para detección de cambios
   */
  createDataHash(data) {
    const hash = crypto.createHash('md5');
    hash.update(JSON.stringify(data, Object.keys(data).sort()));
    return hash.digest('hex');
  }

  // ===== MÉTODOS UTILITARIOS (copiados de Nexus) =====

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

  buildNameParts(input) {
    const transformed = mappingService.transformInputData(input);
    const flat = transformed.__flatIndex || {};

    const firstName = flat['cliente.primer_nombre'] || flat['cliente.nombre'] || input.nombre || '';
    const lastName = flat['cliente.apellido_paterno'] || input.apellido_paterno || input.apellido || '';
    const nombreCliente = `${firstName} ${lastName}`.trim();

    const codigoProspecto =
      input.codigo_de_prospecto ||
      input.codigo_de_cliente ||
      input.codigo ||
      input.id ||
      'SIN_CODIGO';

    return {
      nombreClienteUpper: this.sanitizeForFilenameUpper(nombreCliente),
      codigoProspectoUpper: this.sanitizeForFilenameUpper(codigoProspecto),
    };
  }

  filterNonNullData(data) {
    const filtered = {};
    for (const [key, value] of Object.entries(data)) {
      if (value !== null && value !== undefined && value !== '') {
        filtered[key] = value;
      }
    }
    return filtered;
  }

  getMasterCell(worksheet, addr) {
    const cell = worksheet.getCell(addr);
    if (!cell) return cell;
    if (cell.isMerged && cell.master) {
      const { row, col } = cell.master;
      return worksheet.getCell(row, col);
    }
    if (cell.isMerged && cell.master && cell.master.top && cell.master.left) {
      return worksheet.getCell(cell.master.top, cell.master.left);
    }
    return cell;
  }

  clearFill(cell) {
    cell.fill = {
      type: 'pattern',
      pattern: 'none'
    };
  }

  clearAllRedFillsFromTemplate(worksheet) {
    let cleared = 0;
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.fill &&
            cell.fill.type === 'pattern' &&
            cell.fill.pattern === 'solid' &&
            cell.fill.fgColor &&
            cell.fill.fgColor.argb === 'FFFF0000') {
          this.clearFill(cell);
          cleared++;
        }
      });
    });
    console.log(`[EXCEL-SERVICE] 🧹 Limpiados ${cleared} rellenos rojos de la plantilla`);
  }

  fillCellsFromMappings(worksheet, dataMapping, mappings) {
    let setCount = 0, nulled = 0, skipped = 0;

    console.log('[EXCEL-SERVICE] 🔍 Aplicando mappings...');
    mappings.forEach(({ cell: addr, raw_text }) => {
      const cell = this.getMasterCell(worksheet, addr);
      if (!cell) {
        console.log(`[EXCEL-SERVICE]   ⚠️  Celda ${addr} no encontrada`);
        skipped++;
        return;
      }

      const value = dataMapping.get(raw_text);
      const originalValue = cell.value;

      if (value !== undefined && String(value).trim() !== '') {
        cell.value = value;
        this.clearFill(cell);
        console.log(`[EXCEL-SERVICE]   ✅ ${addr}: "${originalValue}" → "${value}" (${raw_text})`);
        setCount++;
      } else {
        console.log(`[EXCEL-SERVICE]   ⚠️  ${addr}: sin valor para ${raw_text}`);
        cell.value = null;
        nulled++;
      }
    });

    console.log(`[EXCEL-SERVICE] ✍️  Celdas seteadas: ${setCount}, vaciadas: ${nulled}, omitidas: ${skipped}`);
  }

  generateDynamicPassword() {
    const secretPhrase = process.env.FRASE_SECRETA_EXCEL;
    if (!secretPhrase) {
      throw new Error('FRASE_SECRETA_EXCEL no está configurada en las variables de entorno');
    }

    const timestamp = Date.now().toString();
    const hash = crypto.createHash('sha256');
    hash.update(secretPhrase + timestamp);
    const hashHex = hash.digest('hex');
    return hashHex.substring(0, 15);
  }
}

module.exports = new ExcelService();