// =============================================
// SERVICIO DE MAPEO DE PLANTILLAS - SUMATE
// Basado en la implementación de Nexus con adaptaciones
// =============================================

const { storageUtils } = require('../config/supabase');

class MappingService {
  constructor() {
    this.mappings = new Map(); // Mapa de formato -> mappings
    this.placeholderMaps = new Map(); // Mapa de formato -> placeholderMap
    this.cellMaps = new Map(); // Mapa de formato -> cellMap
    this.loaded = new Set(); // Conjunto de formatos cargados

    console.log('[MAPPING-SERVICE] ✅ Servicio inicializado');
  }

  /**
   * Cargar mappings desde Supabase Storage
   */
  async loadMappings(formato = 'general') {
    if (this.loaded.has(formato)) return;

    try {
      let csvFileName;

      // Si el formato incluye extensión .xlsx, limpiarlo
      const cleanFormato = formato.replace('.xlsx', '').replace('.xls', '');

      // Mapeo basado en formato exacto (debe coincidir con excelService)
      switch (cleanFormato) {
        case 'con_HC':
        case 'SCORING_CON_HC':
          csvFileName = 'Mapfield_Con_HC.csv';
          break;
        case 'sin_HC':
        case 'SCORING_SIN_HC':
          csvFileName = 'Mapfield_Sin_HC.csv';
          break;
        case 'seguimiento':
          csvFileName = 'mapfield_seguimiento - Mapfield.csv';
          break;
        case 'evaluacion_economica':
        case 'Evaluacion_Economica_con_Etiquetas':
          csvFileName = 'mapfield_evaluacion_economica.csv';
          break;
        case 'Formato_Editable_Listo':
          csvFileName = 'Mapfield de Placeholders.csv';
          break;
        case 'general':
        default:
          csvFileName = 'Mapfield de Placeholders.csv';
          break;
      }

      console.log(`[MAPPING-SERVICE] 📥 Cargando mapping: ${csvFileName} para formato: ${formato}`);

      // Descargar CSV desde Supabase Storage
      const csvResult = await storageUtils.downloadTemplate(csvFileName);

      if (!csvResult.success) {
        throw new Error(`Error descargando mapping ${csvFileName}: ${csvResult.error}`);
      }

      // Convertir Blob a texto
      const csvText = await csvResult.data.text();

      // Parsear CSV
      const mappings = await this.parseCSV(csvText);

      // Guardar mappings
      this.mappings.set(formato, mappings);
      this.placeholderMaps.set(formato, new Map());
      this.cellMaps.set(formato, new Map());

      // Crear mapas para acceso rápido
      mappings.forEach(mapping => {
        const placeholder = mapping.raw_text; // Con llaves
        const cellAddress = mapping.cell;

        // Mapa placeholder -> celda
        this.placeholderMaps.get(formato).set(placeholder, cellAddress);

        // Mapa celda -> placeholder
        this.cellMaps.get(formato).set(cellAddress, placeholder);
      });

      this.loaded.add(formato);
      console.log(`[MAPPING-SERVICE] ✅ Mappings cargados: ${mappings.length} para formato ${formato}`);

    } catch (error) {
      console.error(`[MAPPING-SERVICE] ❌ Error cargando mappings para formato ${formato}:`, error.message);
      throw error;
    }
  }

  /**
   * Parsear CSV manually (simple parser)
   */
  parseCSV(csvText) {
    return new Promise((resolve) => {
      const lines = csvText.split('\n');
      const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
      const results = [];

      for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;

        const values = line.split(',').map(v => v.trim().replace(/"/g, ''));
        const obj = {};

        headers.forEach((header, index) => {
          obj[header] = values[index] || '';
        });

        if (obj.cell && obj.raw_text) {
          results.push(obj);
        }
      }

      resolve(results);
    });
  }

  /**
   * Obtener mappings para un formato
   */
  getAllMappings(formato = 'general') {
    return this.mappings.get(formato) || [];
  }

  /**
   * Crear mapping de datos
   */
  createDataMapping(inputData, formato = 'general') {
    const transformedData = this.transformInputData(inputData);
    const dataMapping = new Map();
    const mappings = this.mappings.get(formato) || [];

    mappings.forEach(mapping => {
      const placeholderWithBraces = mapping.raw_text; // {campo}
      const path = mapping.placeholder; // campo
      const rawValue = this.extractValueFromPath(transformedData, path);
      const excelValue = this.convertValueForExcel(rawValue);

      dataMapping.set(placeholderWithBraces, excelValue);
    });

    return dataMapping;
  }

  /**
   * Transformar datos de entrada - Adaptado para Sumate
   */
  transformInputData(inputData) {
    const transformed = { ...inputData };

    // Construir índice plano
    const flatIndex = this.flattenObject(transformed);

    // CRÍTICO: Agregar campos planos con guiones bajos al índice con notación de puntos
    // Esto permite que campos como "cliente_primer_nombre" se mapeen a "cliente.primer_nombre"
    for (const [key, value] of Object.entries(inputData)) {
      if (key.includes('_')) {
        const dottedKey = key.replace(/_/g, '.');
        if (flatIndex[dottedKey] === undefined) {
          flatIndex[dottedKey] = value;
        }
      }
    }

    // Alias específicos para Sumate
    const aliasPairs = [
      // Cliente
      ['nombre', 'cliente.nombre'],
      ['apellido', 'cliente.apellido_paterno'],
      ['apellido_paterno', 'cliente.apellido_paterno'],
      ['apellido_materno', 'cliente.apellido_materno'],
      ['telefono', 'cliente.telefono'],
      ['email', 'cliente.email'],
      ['curp', 'cliente.CURP'],
      ['rfc', 'cliente.RFC'],
      ['fecha_nacimiento', 'cliente.fecha_de_nacimiento'],

      // Sumate específicos
      ['credito_solicitado', 'solicitud.monto_solicitado'],
      ['plazo_solicitado', 'solicitud.plazo_solicitado'],
      ['proposito_credito', 'solicitud.proposito'],
      ['score_sumate', 'evaluacion.score_sumate'],
      ['nivel_riesgo', 'evaluacion.nivel_riesgo'],

      // Dirección
      ['calle', 'direccion.calle'],
      ['numero', 'direccion.numero'],
      ['colonia', 'direccion.colonia'],
      ['codigo_postal', 'direccion.codigo_postal'],
      ['municipio', 'direccion.municipio'],
      ['estado', 'direccion.estado'],

      // Trabajo
      ['empresa', 'trabajo.empresa'],
      ['puesto', 'trabajo.puesto'],
      ['salario', 'trabajo.salario_mensual'],
      ['antiguedad', 'trabajo.antiguedad_anos'],

      // Referencias
      ['referencia1_nombre', 'referencias.referencia1.nombre'],
      ['referencia1_telefono', 'referencias.referencia1.telefono'],
      ['referencia2_nombre', 'referencias.referencia2.nombre'],
      ['referencia2_telefono', 'referencias.referencia2.telefono'],

      // Campos comunes planos
      ['id', 'id'],
      ['codigo', 'codigo'],
      ['expediente', 'numero_de_expediente'],
      ['wa_id', 'wa_id'],
      ['fecha', 'fecha'],

      // Alias específicos para buro (CRÍTICO para Scoring)
      ['BC_score', 'buro.BC_score'],
      ['bc_score', 'buro.BC_score'],
      ['calc_bcscore', 'buro.BC_score'], // Para Con HC que usa buro.BC_score
      ['ICC', 'buro.ICC'],
      ['icc', 'buro.ICC'],
      ['no_hit', 'buro.no_hit'],
      ['buro_no_hit', 'buro.no_hit'],

      // calc_bcscore directo (para Sin HC) - CRÍTICO
      ['calc_bcscore', 'calc_bcscore'], // expose directo calc_bcscore

      // Alias para elección final (CRÍTICO para Scoring)
      ['ultima_oferta', 'eleccion_final.monto'],
      ['monto_aceptado', 'eleccion_final.monto'],
      ['monto_aceptado', 'monto_aceptado'], // directo también
      ['calc_capacidad_semanal', 'eleccion_final.pago_semanal'],
      ['pago_semanal', 'eleccion_final.pago_semanal'],

      // Alias para evaluación económica
      ['cuanto_ganas', 'evaluacion_economica.cuanto_ganas'],
      ['cuanto_gastas', 'evaluacion_economica.cuanto_gastas'],
      ['pagos_mensuales_creditos', 'evaluacion_economica.pagos_mensuales_creditos'],
      ['egresos_mensuales', 'evaluacion_economica.egresos_mensuales'],

      // Alias para actividad económica
      ['anos_en_el_negocio', 'actividad_economica.anos_en_el_negocio'],

      // Alias para datos del domicilio
      ['la_casa_es', 'datos_del_domicilio.la_casa_es'],

      // Alias para cliente
      ['edad', 'cliente.edad'],
      ['estado_civil', 'cliente.estado_civil'],
      ['sexo', 'cliente.sexo'],

      // Alias para seguimiento (doble llaves)
      ['nombre_cliente', 'cliente.nombre'],
      ['nombre_asesor', 'nombre_asesor'],
      ['fecha_previo', 'fecha_previo'],
      ['fecha_post', 'fecha_post'],
      ['comentarios_previo', 'comentarios_previo'],
      ['comentarios_post', 'comentarios_post'],

      // Campos específicos de seguimiento (checkboxes sí/no)
      ['monto_cliente_congruente_si', 'seguimiento.monto_cliente_congruente_si'],
      ['monto_cliente_congruente_no', 'seguimiento.monto_cliente_congruente_no'],
      ['riesgo_obligaciones_si', 'seguimiento.riesgo_obligaciones_si'],
      ['riesgo_obligaciones_no', 'seguimiento.riesgo_obligaciones_no'],
      ['riesgo_familiar_credito_si', 'seguimiento.riesgo_familiar_credito_si'],
      ['riesgo_familiar_credito_no', 'seguimiento.riesgo_familiar_credito_no'],
      ['enfermedad_riesgo_credito_si', 'seguimiento.enfermedad_riesgo_credito_si'],
      ['enfermedad_riesgo_credito_no', 'seguimiento.enfermedad_riesgo_credito_no'],
      ['autorizacion_gerente_si', 'seguimiento.autorizacion_gerente_si'],
      ['autorizacion_gerente_no', 'seguimiento.autorizacion_gerente_no'],
      ['problema_funcionamiento_si', 'seguimiento.problema_funcionamiento_si'],
      ['problema_funcionamiento_no', 'seguimiento.problema_funcionamiento_no'],
      ['mismo_aval_si', 'seguimiento.mismo_aval_si'],
      ['mismo_aval_no', 'seguimiento.mismo_aval_no'],
      ['credito_aplicado_si', 'seguimiento.credito_aplicado_si'],
      ['credito_aplicado_no', 'seguimiento.credito_aplicado_no'],
      ['negocio_cambios_si', 'seguimiento.negocio_cambios_si'],
      ['presenta_atrasos_si', 'seguimiento.presenta_atrasos_si'],
      ['presenta_atrasos_no', 'seguimiento.presenta_atrasos_no'],
      ['riesgo_recuperacion_si', 'seguimiento.riesgo_recuperacion_si'],
      ['riesgo_recuperacion_no', 'seguimiento.riesgo_recuperacion_no'],
      ['problema_cliente_si', 'seguimiento.problema_cliente_si'],
      ['problema_cliente_no', 'seguimiento.problema_cliente_no'],

      // Campos de inversión
      ['que_invertir_1', 'seguimiento.que_invertir_1'],
      ['que_invertir_2', 'seguimiento.que_invertir_2'],
      ['que_invertir_3', 'seguimiento.que_invertir_3'],
      ['que_invertir_4', 'seguimiento.que_invertir_4'],
      ['que_invertir_5', 'seguimiento.que_invertir_5'],
      ['valor_estimado_1', 'seguimiento.valor_estimado_1'],
      ['valor_estimado_2', 'seguimiento.valor_estimado_2'],
      ['valor_estimado_3', 'seguimiento.valor_estimado_3'],
      ['valor_estimado_4', 'seguimiento.valor_estimado_4'],
      ['valor_estimado_5', 'seguimiento.valor_estimado_5'],
    ];

    // Aplicar alias
    for (const [fromKey, toKey] of aliasPairs) {
      if (flatIndex[fromKey] !== undefined && flatIndex[toKey] === undefined) {
        flatIndex[toKey] = flatIndex[fromKey];
      }
    }

    // CÁLCULOS AUTOMÁTICOS (como en Nexus)

    // Normalizar calc_bcscore removiendo ceros a la izquierda (CRÍTICO)
    if (flatIndex['calc_bcscore'] !== undefined && flatIndex['calc_bcscore'] !== null) {
      const raw = String(flatIndex['calc_bcscore']).trim();
      const normalized = raw.replace(/^0+/, '');
      flatIndex['calc_bcscore'] = normalized === '' ? '0' : normalized;
    }
    // Fallback: si no hay calc_bcscore, usar bc_score normalizado sin ceros a la izquierda
    if ((flatIndex['calc_bcscore'] === undefined || flatIndex['calc_bcscore'] === '')
        && flatIndex['bc_score'] !== undefined && flatIndex['bc_score'] !== null) {
      const raw = String(flatIndex['bc_score']).trim();
      const normalized = raw.replace(/^0+/, '');
      flatIndex['calc_bcscore'] = normalized === '' ? '0' : normalized;
    }

    // Derivar porcentaje de pago sobre ingreso para sin HC (B6)
    // pct = pagos_mensuales_creditos / cuanto_ganas * 100, formateado con 2 dec y "%"
    const rawIngresos = flatIndex['evaluacion_economica.cuanto_ganas'] ?? flatIndex['cuanto_ganas'];
    const rawPagos = flatIndex['evaluacion_economica.pagos_mensuales_creditos'] ?? flatIndex['pagos_mensuales_creditos'];

    const ingresos = this.parseSmartNumber(rawIngresos);
    const pagos = this.parseSmartNumber(rawPagos);
    if (ingresos > 0 && pagos >= 0) {
      let pct = Math.round(((pagos / ingresos) * 100) * 100) / 100;
      // Si es >0% y <1%, forzar a 1%
      if (pct > 0 && pct < 1) pct = 1;
      flatIndex['evaluacion_economica.pct_pago_sobre_ingreso'] = `${pct}%`;
    }

    // Calcular edad si se proporciona fecha de nacimiento
    try {
      const fecha = flatIndex['cliente.fecha_de_nacimiento'] || flatIndex['cliente_fecha_de_nacimiento'];
      if (fecha && (transformed.cliente?.edad === undefined)) {
        const age = this.calculateAgeFromDateString(String(fecha));
        if (!isNaN(age)) {
          transformed.cliente = { ...(transformed.cliente || {}), edad: age };
          flatIndex['cliente.edad'] = age;
        }
      }
    } catch (_) {
      // ignorar errores de parsing de edad
    }

    // Adjuntar índice plano
    Object.defineProperty(transformed, '__flatIndex', {
      value: flatIndex,
      enumerable: false,
      configurable: false,
      writable: false,
    });

    return transformed;
  }

  /**
   * Aplanar objeto a notación por puntos
   */
  flattenObject(obj, parentKey = '', result = {}) {
    if (!obj || typeof obj !== 'object') return result;

    for (const [key, value] of Object.entries(obj)) {
      const newKey = parentKey ? `${parentKey}.${key}` : key;
      const isPlainObject = value && typeof value === 'object' &&
        !Array.isArray(value) && !(value instanceof Date) && !(Buffer.isBuffer?.(value));

      if (isPlainObject) {
        this.flattenObject(value, newKey, result);
      } else if (Array.isArray(value)) {
        value.forEach((item, idx) => {
          const arrayKey = `${newKey}.${idx}`;
          const isItemObject = item && typeof item === 'object' && !Array.isArray(item);
          if (isItemObject) {
            this.flattenObject(item, arrayKey, result);
          } else {
            result[arrayKey] = item;
          }
        });
      } else {
        result[newKey] = value;
      }
    }

    return result;
  }

  /**
   * Extraer valor de un path
   */
  extractValueFromPath(obj, path) {
    try {
      // Usar índice plano si está disponible
      if (obj && obj.__flatIndex && Object.prototype.hasOwnProperty.call(obj.__flatIndex, path)) {
        return obj.__flatIndex[path];
      }

      // Fallback: buscar versión con guiones bajos
      if (obj && obj.__flatIndex) {
        const underscoreKey = path.replace(/\./g, '_');
        if (Object.prototype.hasOwnProperty.call(obj.__flatIndex, underscoreKey)) {
          return obj.__flatIndex[underscoreKey];
        }
      }

      // Fallback: búsqueda case-insensitive en flatIndex
      if (obj && obj.__flatIndex) {
        const pathLower = path.toLowerCase();
        for (const key in obj.__flatIndex) {
          if (key.toLowerCase() === pathLower) {
            return obj.__flatIndex[key];
          }
        }
      }

      // Navegación anidada tradicional
      return path.split('.').reduce((current, key) => {
        if (current === null || current === undefined) {
          return null;
        }
        return current[key] !== undefined ? current[key] : null;
      }, obj);
    } catch (error) {
      console.error(`[MAPPING-SERVICE] ❌ Error extracting path "${path}":`, error.message);
      return null;
    }
  }

  /**
   * Convertir valor para Excel
   */
  convertValueForExcel(value, placeholder = '') {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    // Si ya es un número, mantenerlo
    if (typeof value === 'number' && !isNaN(value)) {
      return value;
    }

    // Si es una cadena que representa un número
    if (typeof value === 'string') {
      const cleanValue = value.toString().replace(/[$,\s]/g, '');

      if (!isNaN(cleanValue) && cleanValue !== '' && !isNaN(parseFloat(cleanValue))) {
        const numValue = parseFloat(cleanValue);
        return numValue % 1 === 0 ? parseInt(cleanValue) : numValue;
      }
    }

    // Para cualquier otro caso, mantener como string
    return value.toString();
  }

  /**
   * Función para parsear números inteligentemente (como en Nexus)
   */
  parseSmartNumber(v) {
    if (v === null || v === undefined) return NaN;
    const s = String(v).toLowerCase().trim();
    if (!s) return NaN;
    // "35 mil" o "10k"
    if (/\bmil\b/.test(s) || /\bk\b/.test(s)) {
      const n = parseFloat(s.replace(/[^\d.,-]/g, '').replace(',', '.'));
      return isNaN(n) ? NaN : n * 1000;
    }
    const n = parseFloat(s.replace(/[^0-9.-]/g, ''));
    return isNaN(n) ? NaN : n;
  }

  /**
   * Calcular edad desde fecha de nacimiento
   */
  calculateAgeFromDateString(dateStr) {
    try {
      const cleaned = dateStr.replace(/[./]/g, '-');
      let [a, b, c] = cleaned.split('-');

      let yyyy, mm, dd;
      if (a && b && c) {
        if (a.length === 4) {
          // YYYY-MM-DD
          yyyy = parseInt(a, 10); mm = parseInt(b, 10); dd = parseInt(c, 10);
        } else {
          // DD-MM-YYYY
          dd = parseInt(a, 10); mm = parseInt(b, 10); yyyy = parseInt(c, 10);
        }

        const birth = new Date(yyyy, (mm - 1), dd);
        if (isNaN(birth.getTime())) return NaN;

        const today = new Date();
        let age = today.getFullYear() - birth.getFullYear();
        const m = today.getMonth() - birth.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < birth.getDate())) age--;
        return age;
      }

      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return NaN;

      const today = new Date();
      let age = today.getFullYear() - d.getFullYear();
      const m = today.getMonth() - d.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < d.getDate())) age--;
      return age;
    } catch (error) {
      return NaN;
    }
  }

  // Métodos de utilidad
  getPlaceholderForCell(cellAddress, formato = 'general') {
    const cellMap = this.cellMaps.get(formato);
    return cellMap ? cellMap.get(cellAddress) : null;
  }

  getCellForPlaceholder(placeholder, formato = 'general') {
    const placeholderMap = this.placeholderMaps.get(formato);
    return placeholderMap ? placeholderMap.get(placeholder) : null;
  }

  getAllPlaceholders(formato = 'general') {
    const mappings = this.mappings.get(formato) || [];
    return mappings.map(mapping => mapping.raw_text);
  }
}

module.exports = new MappingService();