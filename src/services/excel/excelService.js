// =============================================
// SERVICIO DE GENERACIÓN DE EXCEL - SUMATE
// Basado en la implementación de Nexus con adaptaciones para Sumate
// =============================================

const ExcelJS = require('exceljs');
const axios = require('axios');
const crypto = require('crypto');
const { storageUtils, documentUtils } = require('../../config/supabase');
const mappingService = require('../../utils/mappingService');
const { createLogger, createTemplateTracker } = require('../../utils/logger');

const log = createLogger('EXCEL');

class ExcelService {
  constructor() {
    log.info('Servicio Excel inicializado');
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
        return { success: false, error: 'No hay datos validos para procesar' };
      }

      log.info(`Generando documento formato: ${formato}`);

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
      log.error('Error en generateExcel:', error.message);
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

      log.info(`Procesando webhook para formato: ${formato}`);

      const excelResult = await this.generateExcel(dataSinFormato, formato);

      if (!excelResult.success) {
        return excelResult;
      }

      // Solo subir a storage si se especifica
      let storageUrl = null;
      if (data.saveToStorage === true) {
        const uploadResult = await this.uploadToStorage(excelResult, dataSinFormato);
        if (!uploadResult.success) {
          log.error('Error subiendo a storage:', uploadResult.error);
        } else {
          storageUrl = uploadResult.url;
        }
      }

      // Solo enviar a N8N si está configurado y se solicita
      if (data.sendToN8n !== false && process.env.N8N_WEBHOOK_URL) {
        try {
          await this.sendToN8n(excelResult.fileData, excelResult.fileName, {
            formato: excelResult.formato,
            dataHash: excelResult.dataHash,
            storageUrl: storageUrl
          });
        } catch (error) {
          log.warn('Error enviando a N8N (continuando):', error.message);
        }
      }

      return {
        success: true,
        formato: excelResult.formato,
        fileName: excelResult.fileName,
        fileData: excelResult.fileData,
        buffer: excelResult.buffer,
        storageUrl: storageUrl,
        dataHash: excelResult.dataHash
      };
    } catch (error) {
      log.error('Error procesando webhook:', error.message);
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
        log.warn('Error guardando metadata:', metadataResult.error);
      }

      return {
        success: true,
        url: urlResult.url,
        path: fileName,
        metadata: metadataResult.data
      };
    } catch (error) {
      log.error('Error en uploadToStorage:', error.message);
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
      case 'obligado_solidario':
      case 'obligado':
      case 'ficha_obligado': {
        return `SUMATE_OBLIGADO_SOLIDARIO_${nombreClienteUpper}_${codigoProspectoUpper}_${formattedDate}.xlsx`;
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

      // Mapear formato/template a archivo real y hoja
      switch (formato) {
        case 'con_HC':
        case 'SCORING_CON_HC':
        case 'con_historial_crediticio':
          templateName = 'SCORING_CON_HC.xlsx';
          sheetName = 'Scoring del Cliente';
          break;
        case 'sin_HC':
        case 'SCORING_SIN_HC':
        case 'sin_historial_crediticio':
          templateName = 'SCORING_SIN_HC.xlsx';
          sheetName = 'Scoring del Cliente';
          break;
        case 'obligado_solidario':
        case 'obligado':
        case 'ficha_obligado':
          // Para obligado solidario, usar el Formato Editable que tiene sección de obligado
          templateName = 'Formato_Editable_Listo.xlsx';
          sheetName = 'Obligado Solidario';
          // Si no existe esa hoja, intentar con la principal
          break;
        case 'seguimiento':
          templateName = 'seguimiento del credito (1).xlsx';
          sheetName = 'Hoja1';
          break;
        case 'scoring_etiquetas':
        case 'scoring_con_etiquetas':
          templateName = 'Scoring con Etiquetas.xlsx';
          sheetName = 'Scoring';
          break;
        case 'evaluacion_economica':
        case 'evaluacion':
        case 'ee':
          templateName = 'Evaluacion_Economica_con_Etiquetas.xlsx';
          sheetName = 'EE Simple 2025 ';  // Nombre real con espacio al final
          break;
        case 'Formato_Editable_Listo':
        case 'formato_editable':
        case 'identificacion_cliente':
          templateName = 'Formato_Editable_Listo.xlsx';
          sheetName = 'Ficha de identificación';
          break;
        case 'general':
        default:
          templateName = 'Formato_Editable_Listo.xlsx';
          sheetName = 'Ficha de identificación';
          break;
      }
      log.info(`Usando formato ${formato}: template=${templateName}, hoja=${sheetName}`);

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

      // IMPORTANTE: Forzar recálculo de fórmulas al abrir el archivo
      // Esto evita corrupción del calcChain.xml en plantillas con fórmulas complejas
      workbook.calcProperties = { fullCalcOnLoad: true };

      let worksheet = workbook.getWorksheet(sheetName);

      // Para obligado solidario, si no encuentra la hoja específica, usar la principal
      if (!worksheet && formato.includes('obligado')) {
        log.warn(`Hoja "${sheetName}" no encontrada, usando hoja principal`);
        worksheet = workbook.getWorksheet('Ficha de identificación') ||
                    workbook.getWorksheet('Hoja1') ||
                    workbook.worksheets[0];
      }

      // Para seguimiento, intentar variaciones del nombre de hoja
      if (!worksheet && formato === 'seguimiento') {
        log.warn(`Hoja "${sheetName}" no encontrada para seguimiento, buscando alternativas`);
        worksheet = workbook.getWorksheet('Hoja1') ||
                    workbook.getWorksheet('hoja1') ||
                    workbook.getWorksheet('Sheet1') ||
                    workbook.worksheets[0];
        if (worksheet) {
          log.info(`Usando hoja alternativa: "${worksheet.name}"`);
        }
      }

      if (!worksheet) {
        const availableSheets = workbook.worksheets.map(ws => `"${ws.name}"`);
        log.error(`Hoja "${sheetName}" no encontrada. Disponibles: ${availableSheets.join(', ')}`);
        throw new Error(`Hoja "${sheetName}" no encontrada en la plantilla. Hojas disponibles: ${availableSheets.join(', ')}`);
      }

      log.info(`Plantilla cargada: ${templateName}`);

      // Cargar mappings y aplicar datos
      await mappingService.loadMappings(formato);
      const dataMapping = mappingService.createDataMapping(data, formato);
      const mappings = mappingService.getAllMappings(formato);

      log.info(`Mappings cargados: ${mappings.length} | DataMapping: ${dataMapping.size} entradas`);

      // Crear tracker para logging detallado
      const tracker = createTemplateTracker(formato, log);

      // Limpiar rellenos rojos de la plantilla
      this.clearAllRedFillsFromTemplate(worksheet);

      // Llenar celdas con datos (con tracking)
      this.fillCellsFromMappings(worksheet, dataMapping, mappings, tracker);

      // Imprimir reporte de mapeo
      tracker.printReport();

      // Aplicar reglas específicas por formato
      this.applyFormatSpecificRules(worksheet, formato, data);

      return workbook;
    } catch (error) {
      log.error('Error creando workbook:', error.message);
      throw error;
    }
  }

  /**
   * Aplicar reglas específicas por formato
   */
  applyFormatSpecificRules(worksheet, formato, data) {
    switch (formato) {
      case 'sin_HC':
        // K7/K8 ahora se maneja via mapfield (buro.no_hit.2)
        // No se requiere acción especial
        break;

      case 'expediente_sumate':
        try {
          const password = this.generateDynamicPassword();
          log.info('Aplicando proteccion con contrasena al expediente');

          worksheet.protect(password, {
            selectLockedCells: true,
            selectUnlockedCells: true,
          });

          log.info('Proteccion aplicada exitosamente');
        } catch (protectionError) {
          log.error('Error aplicando proteccion:', protectionError.message);
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

      log.info(`Enviando a N8N: ${fileName}`);

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

      log.info('Enviado exitosamente a N8N');
    } catch (error) {
      log.error('Error enviando a N8N:', error.message);
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
    // Si addr es un rango (ej: "C4:I4"), extraer solo la primera celda
    let cellAddr = addr;
    if (addr && addr.includes(':')) {
      cellAddr = addr.split(':')[0];
    }

    const cell = worksheet.getCell(cellAddr);
    if (!cell) return cell;

    // Debug para celdas problemáticas (K7, K8)
    const isDebugCell = cellAddr === 'K7' || cellAddr === 'K8';
    if (isDebugCell) {
      const masterInfo = cell.master ? `row=${cell.master.row}, col=${cell.master.col}, address=${cell.master.address}` : 'null';
      log.info(`DEBUG CELL ${cellAddr}: isMerged=${cell.isMerged}, master=(${masterInfo}), value="${cell.value}"`);
    }

    // Si la celda está merged, obtener la celda master
    if (cell.isMerged && cell.master) {
      // cell.master puede ser un objeto con row/col o con address
      if (cell.master.row && cell.master.col) {
        const masterCell = worksheet.getCell(cell.master.row, cell.master.col);
        if (isDebugCell) log.info(`DEBUG CELL ${cellAddr}: Redirigiendo a master (${cell.master.row}, ${cell.master.col})`);
        return masterCell;
      }
      if (cell.master.address) {
        const masterCell = worksheet.getCell(cell.master.address);
        if (isDebugCell) log.info(`DEBUG CELL ${cellAddr}: Redirigiendo a master address ${cell.master.address}`);
        return masterCell;
      }
      // Fallback: top/left
      if (cell.master.top && cell.master.left) {
        const masterCell = worksheet.getCell(cell.master.top, cell.master.left);
        if (isDebugCell) log.info(`DEBUG CELL ${cellAddr}: Redirigiendo a master top/left (${cell.master.top}, ${cell.master.left})`);
        return masterCell;
      }
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
    let skippedFormulas = 0;
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        // IMPORTANTE: No tocar celdas que contienen fórmulas para evitar corrupción
        // ExcelJS no soporta bien funciones modernas de Excel 365 (HSTACK, FILTRAR, etc.)
        if (cell.formula || (cell.value && typeof cell.value === 'object' && cell.value.formula)) {
          skippedFormulas++;
          return;
        }

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
    log.debug(`Limpiados ${cleared} rellenos rojos (${skippedFormulas} formulas preservadas)`);
  }

  fillCellsFromMappings(worksheet, dataMapping, mappings, tracker = null) {
    let setCount = 0, nulled = 0, skipped = 0;

    mappings.forEach(({ cell: addr, raw_text, placeholder }) => {
      const cell = this.getMasterCell(worksheet, addr);
      if (!cell) {
        skipped++;
        return;
      }

      let value = dataMapping.get(raw_text);

      if (value !== undefined && String(value).trim() !== '') {
        // Convertir porcentajes: dividir por 100 si es campo de porcentaje
        const isPorcentaje = raw_text.includes('porcentaje_de_ganancia') || raw_text.includes('ingreso_de_ganancia') || raw_text.includes('participacion') || raw_text.includes('pct_pago_sobre_ingreso');
        if (isPorcentaje) {
          const numValue = parseFloat(value);
          if (!isNaN(numValue)) {
            value = numValue / 100;
          }
        }

        cell.value = value;
        this.clearFill(cell);
        setCount++;

        // Registrar en tracker
        if (tracker) {
          tracker.trackMapping(addr, placeholder || raw_text, value, true);
        }
      } else {
        cell.value = '';
        nulled++;

        // Registrar en tracker como vacio
        if (tracker) {
          tracker.trackMapping(addr, placeholder || raw_text, null, false);
        }
      }
    });

    log.info(`Celdas procesadas: ${setCount} colocadas, ${nulled} vacias, ${skipped} omitidas`);
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