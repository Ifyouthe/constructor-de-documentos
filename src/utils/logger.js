// =============================================
// SISTEMA DE LOGGING ESTRUCTURADO - SUMATE
// Logging limpio para debugging de plantillas
// =============================================

const LOG_LEVELS = {
  ERROR: 0,
  WARN: 1,
  INFO: 2,
  DEBUG: 3
};

// Nivel de log configurable via variable de entorno
const currentLevel = LOG_LEVELS[process.env.LOG_LEVEL?.toUpperCase()] ?? LOG_LEVELS.INFO;

class Logger {
  constructor(module) {
    this.module = module;
  }

  _log(level, levelName, message, data = null) {
    if (level > currentLevel) return;

    const timestamp = new Date().toISOString();
    const prefix = `[${timestamp}] [${levelName}] [${this.module}]`;

    if (data !== null) {
      console.log(`${prefix} ${message}`, typeof data === 'object' ? JSON.stringify(data) : data);
    } else {
      console.log(`${prefix} ${message}`);
    }
  }

  error(message, data = null) {
    this._log(LOG_LEVELS.ERROR, 'ERROR', message, data);
  }

  warn(message, data = null) {
    this._log(LOG_LEVELS.WARN, 'WARN', message, data);
  }

  info(message, data = null) {
    this._log(LOG_LEVELS.INFO, 'INFO', message, data);
  }

  debug(message, data = null) {
    this._log(LOG_LEVELS.DEBUG, 'DEBUG', message, data);
  }
}

/**
 * Crear resumen de request para logging
 */
function summarizeRequest(body) {
  return {
    formato: body.formato || body.template || 'general',
    tipo: body.type || 'auto',
    campos_recibidos: Object.keys(body).length,
    campos_principales: Object.keys(body).slice(0, 10),
    tiene_datos_prospecto: !!body.datos_prospecto,
    fichas_solicitadas: body.fichas_a_generar || null
  };
}

/**
 * Clase para trackear el mapeo de datos a plantilla
 */
class TemplateDataTracker {
  constructor(formato, logger) {
    this.formato = formato;
    this.logger = logger;
    this.stats = {
      total_mappings: 0,
      matched: 0,
      not_matched: 0,
      placed: 0,
      empty: 0
    };
    this.details = {
      placed: [],
      empty: [],
      not_found: []
    };
  }

  /**
   * Registrar un mapping procesado
   */
  trackMapping(cellAddr, placeholder, inputValue, wasPlaced) {
    this.stats.total_mappings++;

    const hasValue = inputValue !== undefined && inputValue !== null && String(inputValue).trim() !== '';

    if (hasValue) {
      this.stats.matched++;
      if (wasPlaced) {
        this.stats.placed++;
        this.details.placed.push({ cell: cellAddr, field: placeholder, value: this._truncateValue(inputValue) });
      }
    } else {
      this.stats.not_matched++;
      this.stats.empty++;
      this.details.empty.push({ cell: cellAddr, field: placeholder });
    }
  }

  /**
   * Registrar campo no encontrado en mapfield
   */
  trackNotFound(fieldName) {
    this.details.not_found.push(fieldName);
  }

  /**
   * Truncar valor para mostrar en log
   */
  _truncateValue(value) {
    const str = String(value);
    return str.length > 50 ? str.substring(0, 50) + '...' : str;
  }

  /**
   * Generar reporte de mapeo
   */
  generateReport() {
    return {
      formato: this.formato,
      resumen: {
        total_campos_mapeados: this.stats.total_mappings,
        campos_con_valor: this.stats.matched,
        campos_colocados: this.stats.placed,
        campos_vacios: this.stats.empty,
        porcentaje_exito: this.stats.total_mappings > 0
          ? Math.round((this.stats.placed / this.stats.total_mappings) * 100)
          : 0
      },
      detalles: {
        colocados: this.details.placed.length,
        vacios: this.details.empty.length,
        no_encontrados_en_mapfield: this.details.not_found.length
      }
    };
  }

  /**
   * Imprimir reporte en consola
   */
  printReport() {
    const report = this.generateReport();

    this.logger.info('=== REPORTE DE MAPEO DE PLANTILLA ===');
    this.logger.info(`Formato: ${report.formato}`);
    this.logger.info(`Total campos en mapfield: ${report.resumen.total_campos_mapeados}`);
    this.logger.info(`Campos con valor recibido: ${report.resumen.campos_con_valor}`);
    this.logger.info(`Campos colocados en plantilla: ${report.resumen.campos_colocados}`);
    this.logger.info(`Campos vacios (sin dato): ${report.resumen.campos_vacios}`);
    this.logger.info(`Porcentaje de exito: ${report.resumen.porcentaje_exito}%`);

    if (this.details.not_found.length > 0) {
      this.logger.warn(`Campos en request sin mapfield (${this.details.not_found.length}):`,
        this.details.not_found.slice(0, 10).join(', ') + (this.details.not_found.length > 10 ? '...' : ''));
    }

    // Mostrar primeros 5 campos colocados
    if (this.details.placed.length > 0) {
      this.logger.info('Primeros campos colocados:');
      this.details.placed.slice(0, 5).forEach(item => {
        this.logger.info(`  ${item.cell}: ${item.field} = "${item.value}"`);
      });
      if (this.details.placed.length > 5) {
        this.logger.info(`  ... y ${this.details.placed.length - 5} campos mas`);
      }
    }

    // Mostrar primeros 5 campos vacios
    if (this.details.empty.length > 0 && process.env.LOG_LEVEL === 'DEBUG') {
      this.logger.debug('Campos vacios (primeros 5):');
      this.details.empty.slice(0, 5).forEach(item => {
        this.logger.debug(`  ${item.cell}: ${item.field}`);
      });
    }

    this.logger.info('=== FIN REPORTE ===');

    return report;
  }
}

/**
 * Crear un logger para un modulo especifico
 */
function createLogger(moduleName) {
  return new Logger(moduleName);
}

/**
 * Crear tracker de datos para plantilla
 */
function createTemplateTracker(formato, logger) {
  return new TemplateDataTracker(formato, logger);
}

module.exports = {
  createLogger,
  createTemplateTracker,
  summarizeRequest,
  LOG_LEVELS
};
