// =============================================
// SERVICIO DE GENERACIÓN DE WORD - SUMATE
// Basado en la implementación de Nexus clinicalNoteService
// =============================================

const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const crypto = require('crypto');
const mammoth = require('mammoth');
const { storageUtils, documentUtils } = require('../../config/supabase');
const { createLogger, createTemplateTracker } = require('../../utils/logger');

const log = createLogger('WORD');

class WordService {
  constructor() {
    log.info('Servicio Word inicializado');
  }

  /**
   * Generar documento Word usando plantillas de Supabase Storage
   */
  async generateWord(data, template) {
    try {
      // Filtrar datos no nulos
      const filteredData = this.filterNonNullData(data);

      if (Object.keys(filteredData).length === 0) {
        return { success: false, error: 'No hay datos validos para procesar' };
      }

      log.info(`Generando documento Word: ${template}`);

      // Auto-detectar plantilla Word - no agregar extensión si no la tiene
      let templateName = template;

      // Manejar casos especiales de nombres de plantillas
      if (template === 'obligado_solidario') {
        templateName = 'Fichadeidentificaciondelobligadosolidarioconetiquetas.docx';
      } else if (template === 'ficha_aval') {
        templateName = 'ficha_de_identificacion_del_aval_con_etiquetas.docx';
      } else if (template === 'visita_domiciliaria') {
        templateName = 'Visita domiciliaria con etiquetas.docx';
      } else if (!templateName.includes('.')) {
        // Solo agregar extensión si no tiene ninguna
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
      log.error('Error en generateWord:', error.message);
      return { success: false, error: `Error al generar el archivo Word: ${error.message}` };
    }
  }

  /**
   * Crear documento Word desde plantilla usando docxtemplater (como en Nexus)
   */
  async createDocFromTemplate(data, templateName) {
    try {
      log.info(`Descargando plantilla: ${templateName}`);

      // Descargar plantilla desde Supabase Storage
      const templateResult = await storageUtils.downloadTemplate(templateName);

      if (!templateResult.success) {
        throw new Error(`Error descargando plantilla ${templateName}: ${templateResult.error}`);
      }

      // Convertir Blob a ArrayBuffer y luego a Buffer
      const arrayBuffer = await templateResult.data.arrayBuffer();
      const templateBuffer = Buffer.from(arrayBuffer);

      log.info(`Plantilla cargada: ${templateName}`);

      // Detectar si es archivo .doc (Word 97-2003) y manejarlo diferente
      const isDocFormat = templateName.toLowerCase().endsWith('.doc') && !templateName.toLowerCase().endsWith('.docx');

      if (isDocFormat) {
        log.info('Detectado formato .doc, usando fallback con mammoth');
        return await this.handleDocFormat(templateBuffer, data, templateName);
      }

      // Procesar normalmente para archivos .docx
      return await this.processDocxTemplate(templateBuffer, data, templateName);

    } catch (error) {
      log.error('Error creando documento Word:', error.message);
      throw error;
    }
  }

  /**
   * Procesar plantilla .docx con docxtemplater
   */
  async processDocxTemplate(templateBuffer, data, templateName) {
    try {
      // Cargar plantilla con PizZip
      const zip = new PizZip(templateBuffer);

      // Crear instancia de Docxtemplater
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        nullGetter: function(part) {
          if (!part.module) {
            return '';
          }
          return '';
        }
      });

      // Preparar datos para el template (aplanado)
      const templateData = this.prepareTemplateData(data, templateName);

      // Crear tracker para logging
      const tracker = createTemplateTracker(templateName, log);

      // Contar campos con valor vs vacios
      let camposConValor = 0;
      let camposVacios = 0;
      Object.entries(templateData).forEach(([key, value]) => {
        const hasValue = value !== undefined && value !== null && String(value).trim() !== '';
        if (hasValue) {
          camposConValor++;
          tracker.trackMapping('docx', key, value, true);
        } else {
          camposVacios++;
          tracker.trackMapping('docx', key, null, false);
        }
      });

      log.info(`Datos preparados: ${Object.keys(templateData).length} campos (${camposConValor} con valor, ${camposVacios} vacios)`);

      // Limpiar datos antes de setear en el template para evitar "undefined"
      const cleanedData = this.cleanUndefinedValues(templateData);

      // Renderizar el documento con los datos
      doc.render(cleanedData);

      // Imprimir reporte
      tracker.printReport();

      log.info('Documento Word renderizado exitosamente');

      // Obtener el buffer del documento generado
      const docBuffer = doc.getZip().generate({
        type: 'nodebuffer',
        compression: 'DEFLATE',
      });

      return docBuffer;
    } catch (error) {
      log.error('Error procesando plantilla .docx:', error.message);
      throw error;
    }
  }

  /**
   * Manejar archivos .doc usando mammoth para extraer texto y crear documento simple
   */
  async handleDocFormat(templateBuffer, data, templateName) {
    try {
      log.info('Procesando archivo .doc con mammoth');

      // Extraer texto del archivo .doc
      const result = await mammoth.extractRawText({ buffer: templateBuffer });
      let templateText = result.value;

      log.info(`Texto extraido del .doc, longitud: ${templateText.length}`);

      // Preparar datos para reemplazo
      const templateData = this.prepareTemplateData(data);

      // Limpiar datos antes de procesar placeholders
      const cleanedData = this.cleanUndefinedValues(templateData);

      // Reemplazar placeholders en el texto
      Object.keys(cleanedData).forEach(key => {
        const value = cleanedData[key];
        // Buscar diferentes formatos de placeholder
        const patterns = [
          new RegExp(`\\{\\{${key}\\}\\}`, 'g'),  // {{key}}
          new RegExp(`\\{${key}\\}`, 'g'),       // {key}
          new RegExp(`\\$\\{${key}\\}`, 'g')     // ${key}
        ];

        patterns.forEach(pattern => {
          templateText = templateText.replace(pattern, value);
        });
      });

      // Crear documento Word simple con el texto procesado
      const docContent = this.createSimpleDocx(templateText);

      log.info('Documento .doc procesado exitosamente');
      return docContent;

    } catch (error) {
      log.error('Error procesando archivo .doc:', error.message);
      throw new Error(`Error procesando archivo .doc: ${error.message}`);
    }
  }

  /**
   * Crear documento Word simple con texto
   */
  createSimpleDocx(text) {
    try {
      // Crear estructura básica de un documento .docx
      const zip = new PizZip();

      // Contenido mínimo para un documento Word válido
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>${this.escapeXml(text)}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

      const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

      // Agregar archivos al zip
      zip.file('[Content_Types].xml', contentTypesXml);
      zip.file('_rels/.rels', relsXml);
      zip.file('word/document.xml', documentXml);

      // Generar buffer
      const buffer = zip.generate({
        type: 'nodebuffer',
        compression: 'DEFLATE'
      });

      return buffer;

    } catch (error) {
      log.error('Error creando documento simple:', error.message);
      throw error;
    }
  }

  /**
   * Escapar caracteres especiales para XML
   */
  escapeXml(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /**
   * Preparar datos para el template (como en Nexus)
   */
  prepareTemplateData(data, templateName) {
    // Detectar tipo de template basado en el nombre del template
    const template = templateName || data.template || '';
    const isObligadoSolidario = template.toLowerCase().includes('obligado') ||
                                 template.toLowerCase().includes('fichadeidentificaciondelobligadosolidario');
    const isAval = template.toLowerCase().includes('aval') ||
                   template.toLowerCase().includes('ficha_de_identificacion_del_aval');

    log.debug(`Template: ${template} | ObligadoSolidario: ${isObligadoSolidario} | Aval: ${isAval}`);

    // Si es obligado solidario, usar notación de punto
    if (isObligadoSolidario) {
      return this.prepareObligadoSolidarioData(data);
    }

    // Si es aval, usar notación de punto
    if (isAval) {
      return this.prepareAvalData(data);
    }

    // Para otros documentos, usar estructura plana
    const flatData = this.flattenObject(data);
    const templateData = {};

    // Copiar todos los campos
    Object.keys(flatData).forEach(key => {
      templateData[key] = flatData[key] || '';
    });

    // Agregar campos adicionales comunes
    templateData.fecha = templateData.fecha || new Date().toLocaleDateString();
    templateData.nombre_completo = `${templateData.nombre || ''} ${templateData.apellido_paterno || templateData.apellido || ''}`.trim();

    return templateData;
  }

  /**
   * Preparar datos específicos para Obligado Solidario
   */
  prepareObligadoSolidarioData(data) {
    log.debug('Preparando datos para obligado solidario');

    // Crear objeto con notación de punto que espera docxtemplater
    const templateData = {
      // Campos de nivel superior
      'codigo': data.codigo || data.codigo_de_prospecto || '',
      'fecha': data.fecha || new Date().toLocaleDateString('es-MX')
    };

    // Si los datos ya vienen con notación de punto, usarlos directamente
    if (data['obligado.primer_nombre'] !== undefined) {
      // Copiar todos los campos con notación de punto
      Object.keys(data).forEach(key => {
        if (key.includes('.') || key === 'codigo' || key === 'fecha') {
          templateData[key] = data[key] || '';
        }
      });

      return templateData;
    }

    // Si los datos vienen con estructura anidada, convertir a notación de punto
    if (data.obligado && typeof data.obligado === 'object') {

      // Convertir obligado
      const obligado = data.obligado;
      templateData['obligado.primer_nombre'] = obligado.primer_nombre || '';
      templateData['obligado.segundo_nombre'] = obligado.segundo_nombre || '';
      templateData['obligado.apellido_paterno'] = obligado.apellido_paterno || '';
      templateData['obligado.apellido_materno'] = obligado.apellido_materno || '';
      templateData['obligado.clave_de_elector'] = obligado.clave_de_elector || '';
      templateData['obligado.CURP'] = obligado.CURP || obligado.curp || '';
      templateData['obligado.RFC'] = obligado.RFC || obligado.rfc || '';
      templateData['obligado.firma_electronica'] = obligado.firma_electronica || '';
      templateData['obligado.nacionalidad'] = obligado.nacionalidad || '';
      templateData['obligado.pais_de_nacimiento'] = obligado.pais_de_nacimiento || '';
      templateData['obligado.estado_de_nacimiento'] = obligado.estado_de_nacimiento || '';
      templateData['obligado.fecha_de_nacimiento'] = obligado.fecha_de_nacimiento || '';
      templateData['obligado.estado_civil'] = obligado.estado_civil || '';
      templateData['obligado.depentientes_economicos'] = obligado.depentientes_economicos || obligado.dependientes_economicos || '';
      templateData['obligado.sexo'] = obligado.sexo || '';
      templateData['obligado.escolaridad'] = obligado.escolaridad || '';
      templateData['obligado.actividad'] = obligado.actividad || '';
      templateData['obligado.profesion'] = obligado.profesion || '';
      templateData['obligado.ocupacion'] = obligado.ocupacion || '';

      // Convertir domicilio
      const domicilio = data.domicilio || {};
      templateData['domicilio.direccion_calle'] = domicilio.direccion_calle || '';
      templateData['domicilio.direccion_numero'] = domicilio.direccion_numero || '';
      templateData['domicilio.direccion_colonia'] = domicilio.direccion_colonia || '';
      templateData['domicilio.direccion_ciudad'] = domicilio.direccion_ciudad || '';
      templateData['domicilio.codigo_postal'] = domicilio.codigo_postal || '';
      templateData['domicilio.municipio'] = domicilio.municipio || '';
      templateData['domicilio.estado'] = domicilio.estado || '';
      templateData['domicilio.pais'] = domicilio.pais || '';
      templateData['domicilio.referecia_localizacion'] = domicilio.referecia_localizacion || domicilio.referencia_localizacion || '';
      templateData['domicilio.la_casa_es'] = domicilio.la_casa_es || '';
      templateData['domicilio.telefono'] = domicilio.telefono || '';

      // Convertir cargo_publico (con notación anidada correcta para familiares)
      const cargo_publico = data.cargo_publico || {};
      templateData['cargo_publico.si'] = cargo_publico.si || '';
      templateData['cargo_publico.no'] = cargo_publico.no || '';
      templateData['cargo_publico.familiares.si'] = cargo_publico.familiares?.si || data.cargo_publico_familiares?.si || '';
      templateData['cargo_publico.familiares.no'] = cargo_publico.familiares?.no || data.cargo_publico_familiares?.no || '';

      // Convertir protesta
      const protesta = data.protesta || {};
      templateData['protesta.es_accionista'] = protesta.es_accionista || '';
      templateData['protesta.tiene_relacion_con_accionista'] = protesta.tiene_relacion_con_accionista || '';

      return templateData;
    }

    // Si los datos vienen planos, crear notación de punto directamente

    // Campos obligado con notación de punto
    templateData['obligado.primer_nombre'] = data.primer_nombre || '';
    templateData['obligado.segundo_nombre'] = data.segundo_nombre || '';
    templateData['obligado.apellido_paterno'] = data.apellido_paterno || '';
    templateData['obligado.apellido_materno'] = data.apellido_materno || '';
    templateData['obligado.clave_de_elector'] = data.clave_de_elector || '';
    templateData['obligado.CURP'] = data.CURP || data.curp || '';
    templateData['obligado.RFC'] = data.RFC || data.rfc || '';
    templateData['obligado.firma_electronica'] = data.firma_electronica || '';
    templateData['obligado.nacionalidad'] = data.nacionalidad || '';
    templateData['obligado.pais_de_nacimiento'] = data.pais_de_nacimiento || '';
    templateData['obligado.estado_de_nacimiento'] = data.estado_de_nacimiento || '';
    templateData['obligado.fecha_de_nacimiento'] = data.fecha_de_nacimiento || '';
    templateData['obligado.estado_civil'] = data.estado_civil || '';
    templateData['obligado.depentientes_economicos'] = data.depentientes_economicos || data.dependientes_economicos || '';
    templateData['obligado.sexo'] = data.sexo || '';
    templateData['obligado.escolaridad'] = data.escolaridad || '';
    templateData['obligado.actividad'] = data.actividad || '';
    templateData['obligado.profesion'] = data.profesion || '';
    templateData['obligado.ocupacion'] = data.ocupacion || '';

    // Campos domicilio con notación de punto
    templateData['domicilio.direccion_calle'] = data.direccion_calle || '';
    templateData['domicilio.direccion_numero'] = data.direccion_numero || '';
    templateData['domicilio.direccion_colonia'] = data.direccion_colonia || '';
    templateData['domicilio.direccion_ciudad'] = data.direccion_ciudad || '';
    templateData['domicilio.codigo_postal'] = data.codigo_postal || '';
    templateData['domicilio.municipio'] = data.municipio || '';
    templateData['domicilio.estado'] = data.estado || '';
    templateData['domicilio.pais'] = data.pais || '';
    templateData['domicilio.referecia_localizacion'] = data.referecia_localizacion || data.referencia_localizacion || '';
    templateData['domicilio.la_casa_es'] = data.la_casa_es || '';
    templateData['domicilio.telefono'] = data.telefono || '';

    // Campos cargo_publico con notación de punto
    templateData['cargo_publico.si'] = data.cargo_publico_si || '';
    templateData['cargo_publico.no'] = data.cargo_publico_no || '';
    templateData['cargo_publico.familiares.si'] = data.cargo_publico_familiares_si || '';
    templateData['cargo_publico.familiares.no'] = data.cargo_publico_familiares_no || '';

    // Campos protesta con notación de punto
    templateData['protesta.es_accionista'] = data.es_accionista || '';
    templateData['protesta.tiene_relacion_con_accionista'] = data.tiene_relacion_con_accionista || '';

    return templateData;
  }

  /**
   * Preparar datos específicos para Aval
   */
  prepareAvalData(data) {
    log.debug('Preparando datos para aval');

    // Crear objeto con notación de punto que espera docxtemplater
    const templateData = {
      // Campos de nivel superior
      'codigo': data.codigo || data.codigo_de_prospecto || '',
      'fecha': data.fecha || new Date().toLocaleDateString('es-MX')
    };

    // Si los datos ya vienen con notación de punto, usarlos directamente
    if (data['aval.primer_nombre'] !== undefined) {
      // Copiar todos los campos con notación de punto
      Object.keys(data).forEach(key => {
        if (key.includes('.') || key === 'codigo' || key === 'fecha') {
          templateData[key] = data[key] || '';
        }
      });

      return templateData;
    }

    // Si los datos vienen con estructura anidada, convertir a notación de punto
    if (data.aval && typeof data.aval === 'object') {

      // Convertir aval
      const aval = data.aval;
      templateData['aval.primer_nombre'] = aval.primer_nombre || '';
      templateData['aval.segundo_nombre'] = aval.segundo_nombre || '';
      templateData['aval.apellido_paterno'] = aval.apellido_paterno || '';
      templateData['aval.apellido_materno'] = aval.apellido_materno || '';
      templateData['aval.clave_de_elector'] = aval.clave_de_elector || '';
      templateData['aval.CURP'] = aval.CURP || aval.curp || '';
      templateData['aval.RFC'] = aval.RFC || aval.rfc || '';
      templateData['aval.firma_electronica'] = aval.firma_electronica || '';
      templateData['aval.nacionalidad'] = aval.nacionalidad || '';
      templateData['aval.pais_de_nacimiento'] = aval.pais_de_nacimiento || '';
      templateData['aval.estado_de_nacimiento'] = aval.estado_de_nacimiento || '';
      templateData['aval.fecha_de_nacimiento'] = aval.fecha_de_nacimiento || '';
      templateData['aval.estado_civil'] = aval.estado_civil || '';
      templateData['aval.depentientes_economicos'] = aval.depentientes_economicos || aval.dependientes_economicos || '';
      templateData['aval.sexo'] = aval.sexo || '';
      templateData['aval.escolaridad'] = aval.escolaridad || '';
      templateData['aval.actividad'] = aval.actividad || '';
      templateData['aval.profesion'] = aval.profesion || '';
      templateData['aval.ocupacion'] = aval.ocupacion || '';

      // Convertir domicilio
      const domicilio = data.domicilio || {};
      templateData['domicilio.direccion_calle'] = domicilio.direccion_calle || '';
      templateData['domicilio.direccion_numero'] = domicilio.direccion_numero || '';
      templateData['domicilio.direccion_colonia'] = domicilio.direccion_colonia || '';
      templateData['domicilio.direccion_ciudad'] = domicilio.direccion_ciudad || '';
      templateData['domicilio.codigo_postal'] = domicilio.codigo_postal || '';
      templateData['domicilio.municipio'] = domicilio.municipio || '';
      templateData['domicilio.estado'] = domicilio.estado || '';
      templateData['domicilio.pais'] = domicilio.pais || '';
      templateData['domicilio.referecia_localizacion'] = domicilio.referecia_localizacion || domicilio.referencia_localizacion || '';
      templateData['domicilio.la_casa_es'] = domicilio.la_casa_es || '';
      templateData['domicilio.telefono'] = domicilio.telefono || '';

      // Convertir pareja (si existe)
      const pareja = data.pareja || {};
      templateData['pareja.apellido_paterno'] = pareja.apellido_paterno || '';
      templateData['pareja.apellido_materno'] = pareja.apellido_materno || '';
      templateData['pareja.primer_nombre'] = pareja.primer_nombre || '';
      templateData['pareja.segundo_nombre'] = pareja.segundo_nombre || '';
      templateData['pareja.estado_de_nacimiento'] = pareja.estado_de_nacimiento || '';
      templateData['pareja.fecha_de_nacimiento'] = pareja.fecha_de_nacimiento || '';
      templateData['pareja.ocupacion'] = pareja.ocupacion || '';
      templateData['pareja.lugar_de_trabajo'] = pareja.lugar_de_trabajo || '';
      templateData['pareja.clave_de_elector'] = pareja.clave_de_elector || '';
      templateData['pareja.CURP'] = pareja.CURP || pareja.curp || '';
      templateData['pareja.escolaridad'] = pareja.escolaridad || '';

      // Convertir cargo_publico (con notación anidada correcta para familiares)
      const cargo_publico = data.cargo_publico || {};
      templateData['cargo_publico.si'] = cargo_publico.si || '';
      templateData['cargo_publico.no'] = cargo_publico.no || '';
      templateData['cargo_publico.familiares.si'] = cargo_publico.familiares?.si || data.cargo_publico_familiares?.si || '';
      templateData['cargo_publico.familiares.no'] = cargo_publico.familiares?.no || data.cargo_publico_familiares?.no || '';

      // Convertir protesta
      const protesta = data.protesta || {};
      templateData['protesta.es_accionista'] = protesta.es_accionista || '';
      templateData['protesta.tiene_relacion_con_accionista'] = protesta.tiene_relacion_con_accionista || '';

      // Convertir asalariado (si existe)
      const asalariado = data.asalariado || {};
      templateData['asalariado.nombre_de_la_empresa'] = asalariado.nombre_de_la_empresa || '';
      templateData['asalariado.ubicacion'] = asalariado.ubicacion || '';
      templateData['asalariado.puesto'] = asalariado.puesto || '';
      templateData['asalariado.calle'] = asalariado.calle || '';
      templateData['asalariado.numero'] = asalariado.numero || '';
      templateData['asalariado.colonia'] = asalariado.colonia || '';
      templateData['asalariado.ciudad'] = asalariado.ciudad || '';
      templateData['asalariado.codigo_postal'] = asalariado.codigo_postal || '';
      templateData['asalariado.municipio'] = asalariado.municipio || '';
      templateData['asalariado.estado'] = asalariado.estado || '';
      templateData['asalariado.pais'] = asalariado.pais || '';
      templateData['asalariado.telefono_de_la_empresa'] = asalariado.telefono_de_la_empresa || '';
      templateData['asalariado.dias_trabajo'] = asalariado.dias_trabajo || '';
      templateData['asalariado.dias_descanso'] = asalariado.dias_descanso || '';
      templateData['asalariado.sueldo'] = asalariado.sueldo || '';
      templateData['asalariado.otros_ingresos'] = asalariado.otros_ingresos || '';
      templateData['asalariado.gastos'] = asalariado.gastos || '';
      templateData['asalariado.ingreso_disponible'] = asalariado.ingreso_disponible || '';

      // Convertir microempresario (si existe)
      const microempresario = data.microempresario || {};
      templateData['microempresario.negocio'] = microempresario.negocio || '';
      templateData['microempresario.sector'] = microempresario.sector || '';
      templateData['microempresario.tipo_de_negocio'] = microempresario.tipo_de_negocio || '';
      templateData['microempresario.ubicacion'] = microempresario.ubicacion || '';
      templateData['microempresario.local'] = microempresario.local || '';
      templateData['microempresario.direccion.calle'] = microempresario.direccion?.calle || '';
      templateData['microempresario.direccion.numero'] = microempresario.direccion?.numero || '';
      templateData['microempresario.direccion.colonia'] = microempresario.direccion?.colonia || '';
      templateData['microempresario.direccion.ciudad'] = microempresario.direccion?.ciudad || '';
      templateData['microempresario.direccion.codigo_postal'] = microempresario.direccion?.codigo_postal || '';
      templateData['microempresario.direccion.municipio'] = microempresario.direccion?.municipio || '';
      templateData['microempresario.direccion.estado'] = microempresario.direccion?.estado || '';
      templateData['microempresario.direccion.pais'] = microempresario.direccion?.pais || '';
      templateData['microempresario.telefono_del_negocio'] = microempresario.telefono_del_negocio || '';
      templateData['microempresario.años_en_el_negocio'] = microempresario.años_en_el_negocio || '';
      templateData['microempresario.dias_trabajo'] = microempresario.dias_trabajo || '';
      templateData['microempresario.dias_descanso'] = microempresario.dias_descanso || '';
      templateData['microempresario.horario_de_trabajo'] = microempresario.horario_de_trabajo || '';
      templateData['microempresario.otro_ingreso_o_negocio'] = microempresario.otro_ingreso_o_negocio || '';
      templateData['microempresario.otro_ingreso_o_negocio.cual'] = microempresario.otro_ingreso_o_negocio?.cual || '';

      return templateData;
    }

    // Si los datos vienen planos, crear notación de punto directamente

    // Campos aval con notación de punto
    templateData['aval.primer_nombre'] = data.primer_nombre || '';
    templateData['aval.segundo_nombre'] = data.segundo_nombre || '';
    templateData['aval.apellido_paterno'] = data.apellido_paterno || '';
    templateData['aval.apellido_materno'] = data.apellido_materno || '';
    templateData['aval.clave_de_elector'] = data.clave_de_elector || '';
    templateData['aval.CURP'] = data.CURP || data.curp || '';
    templateData['aval.RFC'] = data.RFC || data.rfc || '';
    templateData['aval.firma_electronica'] = data.firma_electronica || '';
    templateData['aval.nacionalidad'] = data.nacionalidad || '';
    templateData['aval.pais_de_nacimiento'] = data.pais_de_nacimiento || '';
    templateData['aval.estado_de_nacimiento'] = data.estado_de_nacimiento || '';
    templateData['aval.fecha_de_nacimiento'] = data.fecha_de_nacimiento || '';
    templateData['aval.estado_civil'] = data.estado_civil || '';
    templateData['aval.depentientes_economicos'] = data.depentientes_economicos || data.dependientes_economicos || '';
    templateData['aval.sexo'] = data.sexo || '';
    templateData['aval.escolaridad'] = data.escolaridad || '';
    templateData['aval.actividad'] = data.actividad || '';
    templateData['aval.profesion'] = data.profesion || '';
    templateData['aval.ocupacion'] = data.ocupacion || '';

    // Campos domicilio con notación de punto
    templateData['domicilio.direccion_calle'] = data.direccion_calle || '';
    templateData['domicilio.direccion_numero'] = data.direccion_numero || '';
    templateData['domicilio.direccion_colonia'] = data.direccion_colonia || '';
    templateData['domicilio.direccion_ciudad'] = data.direccion_ciudad || '';
    templateData['domicilio.codigo_postal'] = data.codigo_postal || '';
    templateData['domicilio.municipio'] = data.municipio || '';
    templateData['domicilio.estado'] = data.estado || '';
    templateData['domicilio.pais'] = data.pais || '';
    templateData['domicilio.referecia_localizacion'] = data.referecia_localizacion || data.referencia_localizacion || '';
    templateData['domicilio.la_casa_es'] = data.la_casa_es || '';
    templateData['domicilio.telefono'] = data.telefono || '';

    // Campos pareja con notación de punto (si vienen planos con prefijo pareja_)
    templateData['pareja.apellido_paterno'] = data.pareja_apellido_paterno || '';
    templateData['pareja.apellido_materno'] = data.pareja_apellido_materno || '';
    templateData['pareja.primer_nombre'] = data.pareja_primer_nombre || '';
    templateData['pareja.segundo_nombre'] = data.pareja_segundo_nombre || '';
    templateData['pareja.estado_de_nacimiento'] = data.pareja_estado_de_nacimiento || '';
    templateData['pareja.fecha_de_nacimiento'] = data.pareja_fecha_de_nacimiento || '';
    templateData['pareja.ocupacion'] = data.pareja_ocupacion || '';
    templateData['pareja.lugar_de_trabajo'] = data.pareja_lugar_de_trabajo || '';
    templateData['pareja.clave_de_elector'] = data.pareja_clave_de_elector || '';
    templateData['pareja.CURP'] = data.pareja_CURP || data.pareja_curp || '';
    templateData['pareja.escolaridad'] = data.pareja_escolaridad || '';

    // Campos cargo_publico con notación de punto
    templateData['cargo_publico.si'] = data.cargo_publico_si || '';
    templateData['cargo_publico.no'] = data.cargo_publico_no || '';
    templateData['cargo_publico.familiares.si'] = data.cargo_publico_familiares_si || '';
    templateData['cargo_publico.familiares.no'] = data.cargo_publico_familiares_no || '';

    // Campos protesta con notación de punto
    templateData['protesta.es_accionista'] = data.es_accionista || '';
    templateData['protesta.tiene_relacion_con_accionista'] = data.tiene_relacion_con_accionista || '';

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

      log.info(`Procesando webhook para template: ${template}`);

      const wordResult = await this.generateWord(dataSinTemplate, template);

      if (!wordResult.success) {
        return wordResult;
      }

      // Solo subir a storage si se especifica
      let storageUrl = null;
      if (data.saveToStorage === true) {
        const uploadResult = await this.uploadToStorage(wordResult, dataSinTemplate);
        if (!uploadResult.success) {
          log.error('Error subiendo a storage:', uploadResult.error);
        } else {
          storageUrl = uploadResult.url;
        }
      }

      // N8N ya no se usa - el envío se maneja desde el servicio externo que llama a este webhook

      return {
        success: true,
        template: wordResult.template,
        fileName: wordResult.fileName,
        fileData: wordResult.fileData,
        buffer: wordResult.buffer,
        storageUrl: storageUrl,
        dataHash: wordResult.dataHash
      };
    } catch (error) {
      log.error('Error procesando webhook:', error.message);
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
   * Limpiar valores undefined/null/vacíos para evitar que aparezca "undefined" en los documentos
   */
  cleanUndefinedValues(data) {
    const cleaned = {};

    Object.keys(data).forEach(key => {
      const value = data[key];

      // Si es undefined, null, o la string "undefined", convertir a string vacío
      if (value === undefined || value === null || value === 'undefined' || value === '') {
        cleaned[key] = '';
      } else {
        // Convertir a string y limpiar
        const stringValue = String(value);
        cleaned[key] = stringValue === 'undefined' ? '' : stringValue;
      }
    });

    return cleaned;
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
    const nombre = data.nombre || data.nombre_del_cliente || data.primer_nombre || data['cliente.primer_nombre'] || 'SIN_NOMBRE';
    const apellido = data.apellido_paterno || data.apellido || data['cliente.apellido_paterno'] || '';
    const codigo = data.codigo || data.codigo_de_prospecto || data.id || 'SIN_CODIGO';

    const nombreCompleto = `${nombre} ${apellido}`.trim();
    const nombreSanitized = this.sanitizeForFilenameUpper(nombreCompleto);
    const codigoSanitized = this.sanitizeForFilenameUpper(codigo);

    const extension = templateName.toLowerCase().endsWith('.docx') ? '.docx' : '.doc';

    return `SUMATE_${baseName}_${nombreSanitized}_${codigoSanitized}_${formattedDate}${extension}`;
  }

  /**
   * Filtrar datos no nulos - convertir undefined/null a string vacío en lugar de eliminar
   */
  filterNonNullData(data) {
    const filtered = {};
    for (const [key, value] of Object.entries(data)) {
      // En lugar de eliminar campos undefined/null, convertirlos a string vacío
      if (value === null || value === undefined || value === 'undefined') {
        filtered[key] = '';
      } else if (value === '') {
        filtered[key] = '';
      } else {
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