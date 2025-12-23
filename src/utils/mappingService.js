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
      // Auto-detectar CSV basado en el formato (como en Nexus)
      let csvFileName = `Mapfield_${formato}.csv`;

      // Si es 'general', usar el primer CSV disponible como fallback
      if (formato === 'general') {
        // Intentar listar CSVs disponibles
        const templatesResult = await storageUtils.listTemplates();
        if (templatesResult.success && templatesResult.templates.length > 0) {
          const csvFiles = templatesResult.templates.filter(t => t.name.endsWith('.csv'));
          if (csvFiles.length > 0) {
            csvFileName = csvFiles[0].name;
          }
        }
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
    ];

    // Aplicar alias
    for (const [fromKey, toKey] of aliasPairs) {
      if (flatIndex[fromKey] !== undefined && flatIndex[toKey] === undefined) {
        flatIndex[toKey] = flatIndex[fromKey];
      }
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