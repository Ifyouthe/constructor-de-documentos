// =============================================
// SERVICIO DE GENERACIÓN MÚLTIPLE DE DOCUMENTOS - SUMATE
// Permite generar múltiples fichas en una sola request
// =============================================

const excelService = require('./excel/excelService');
const wordService = require('./word/wordService');
const mappingService = require('../utils/mappingService');

class MultipleDocumentsService {
  constructor() {
    // Definición de mapeos por tipo de ficha basado en los normalizadores de N8N
    this.fichaMappers = {
      identificacion_cliente: this.mapIdentificacionCliente.bind(this),
      visita_domiciliaria: this.mapVisitaDomiciliaria.bind(this),
      evaluacion_economica_simple: this.mapEvaluacionEconomicaSimple.bind(this),
      obligado_solidario: this.mapObligadoSolidario.bind(this),
      aval: this.mapAval.bind(this),
      seguimiento_previo: this.mapSeguimientoPrevio.bind(this),
      scoring_con_hc: this.mapScoringConHC.bind(this),
      scoring_sin_hc: this.mapScoringSinHC.bind(this),
      seguimiento_credito: this.mapSeguimientoCredito.bind(this)
    };

    console.log('[MULTIPLE-DOCS-SERVICE] ✅ Servicio inicializado');
  }

  /**
   * Generar múltiples documentos en una sola request
   */
  async generateMultipleDocuments(fichas_a_generar, datos_prospecto) {
    try {
      console.log('[MULTIPLE-DOCS] 📨 Generando múltiples fichas:', fichas_a_generar);

      const resultados = {
        success: true,
        documentos_generados: [],
        errores: [],
        metadata: {
          total_solicitados: fichas_a_generar.length,
          total_generados: 0,
          total_errores: 0,
          timestamp: new Date().toISOString()
        }
      };

      // Procesar cada ficha
      for (const tipoFicha of fichas_a_generar) {
        try {
          console.log(`[MULTIPLE-DOCS] 🔄 Procesando ficha: ${tipoFicha}`);

          const resultado = await this.generateSingleDocument(tipoFicha, datos_prospecto);

          if (resultado.success) {
            resultados.documentos_generados.push(resultado);
            resultados.metadata.total_generados++;
            console.log(`[MULTIPLE-DOCS] ✅ Ficha generada: ${tipoFicha}`);
          } else {
            resultados.errores.push({
              tipo_ficha: tipoFicha,
              error: resultado.error
            });
            resultados.metadata.total_errores++;
            console.error(`[MULTIPLE-DOCS] ❌ Error en ficha ${tipoFicha}:`, resultado.error);
          }
        } catch (error) {
          resultados.errores.push({
            tipo_ficha: tipoFicha,
            error: error.message
          });
          resultados.metadata.total_errores++;
          console.error(`[MULTIPLE-DOCS] ❌ Error crítico en ficha ${tipoFicha}:`, error.message);
        }
      }

      // Si no se generó ningún documento, marcar como error
      if (resultados.metadata.total_generados === 0) {
        resultados.success = false;
      }

      return resultados;

    } catch (error) {
      console.error('[MULTIPLE-DOCS] ❌ Error crítico:', error.message);
      return {
        success: false,
        error: 'Error crítico en generación múltiple',
        detalles: error.message,
        metadata: {
          timestamp: new Date().toISOString()
        }
      };
    }
  }

  /**
   * Generar un documento individual
   */
  async generateSingleDocument(tipoFicha, datosProspecto) {
    try {
      // Verificar si tenemos mapper para esta ficha
      if (!this.fichaMappers[tipoFicha]) {
        throw new Error(`Tipo de ficha no soportado: ${tipoFicha}`);
      }

      // Mapear datos según el tipo de ficha
      const datosNormalizados = this.fichaMappers[tipoFicha](datosProspecto);

      let resultado;

      // Determinar servicio y generar documento
      switch (tipoFicha) {
        case 'identificacion_cliente':
        case 'evaluacion_economica_simple':
        case 'scoring_con_hc':
        case 'scoring_sin_hc':
        case 'seguimiento_credito':
        case 'seguimiento_previo':
          resultado = await excelService.generateExcel(datosNormalizados.data, datosNormalizados.template);
          break;

        case 'visita_domiciliaria':
        case 'obligado_solidario':
        case 'aval':
          resultado = await wordService.generateWord(datosNormalizados.data, datosNormalizados.template);
          break;

        default:
          throw new Error(`Servicio no definido para ficha: ${tipoFicha}`);
      }

      if (!resultado.success) {
        throw new Error(resultado.error);
      }

      // Crear nombre de archivo con formato específico SOLO para múltiples documentos
      const customFileName = this.buildCustomFileName(datosProspecto, tipoFicha, resultado.fileName);

      return {
        success: true,
        tipo_ficha: tipoFicha,
        fileName: customFileName,
        fileData: resultado.fileData,
        formato: resultado.formato || tipoFicha,
        metadata: {
          template_usado: datosNormalizados.template,
          timestamp: new Date().toISOString()
        }
      };

    } catch (error) {
      return {
        success: false,
        tipo_ficha: tipoFicha,
        error: error.message
      };
    }
  }

  /**
   * MAPEOS POR TIPO DE FICHA
   * Basados en los normalizadores de N8N que me compartiste
   */

  /**
   * Mapeo para Identificación del Cliente (Excel)
   */
  mapIdentificacionCliente(datos) {
    // Helper functions del mapeo original
    const empty = (v) => (v === null || v === undefined ? "" : v);
    const boolX = (v) => v === true || v === "true" || v === "X" ? "X" : "";

    // Fecha hoy DD/MM/YYYY
    const today = () => {
      const d = new Date();
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const yyyy = d.getFullYear();
      return `${dd}/${mm}/${yyyy}`;
    };

    // Mapeo según el código de N8N
    const mappedData = {
      // Identificación
      codigo_de_prospecto: empty(datos.codigo_de_prospecto) || empty(datos.id_expediente),
      codigo_de_cliente: empty(datos.codigo_de_cliente),
      fecha_identificacion: empty(datos.fecha_identificacion_cliente) || today(),

      // Cliente - mapear desde datos_prospecto
      cliente_primer_nombre: empty(datos.primer_nombre),
      cliente_segundo_nombre: empty(datos.segundo_nombre),
      cliente_apellido_paterno: empty(datos.primer_apellido),
      cliente_apellido_materno: empty(datos.segundo_apellido),
      cliente_fecha_de_nacimiento: empty(datos.fecha_nacimiento),
      cliente_sexo: empty(datos.sexo),
      cliente_curp: empty(datos.curp || datos.cedula),
      cliente_rfc: empty(datos.rfc),
      cliente_escolaridad: empty(datos.escolaridad),
      cliente_correo_electronico: empty(datos.correo),
      cliente_clave_de_elector: empty(datos.clave_de_elector),
      cliente_nacionalidad: empty(datos.nacionalidad),
      cliente_estado_de_nacimiento: empty(datos.estado_nacimiento),
      cliente_pais_de_nacimiento: empty(datos.direccion_pais),
      cliente_profesion: empty(datos.profesion),
      cliente_dependientes_economicos: empty(datos.dependientes_economicos),
      estado_civil: empty(datos.estado_civil),
      cliente_uso_de_redes_sociales: empty(datos.uso_redes_sociales),
      cliente_uso_de_redes_sociales_cual: empty(datos.redes_sociales_cual),
      cliente_uso_de_redes_sociales_usuario: empty(datos.usuario_redes_sociales),
      cliente_efirma: empty(datos.efirma),
      cliente_efirma_si: boolX(datos.efirma_si),
      cliente_efirma_no: boolX(datos.efirma_no),

      // Domicilio
      datos_del_domicilio_telefono: empty(datos.telefono),
      datos_del_domicilio_pais: empty(datos.direccion_pais),
      datos_del_domicilio_estado: empty(datos.direccion_provincia),
      datos_del_domicilio_localidad: empty(datos.direccion_ciudad),
      datos_del_domicilio_municipio: empty(datos.municipio),
      datos_del_domicilio_la_casa_es: empty(datos.la_casa_es),
      datos_del_domicilio_referencia_de_localizacion: empty(datos.referencia_localizacion),
      datos_del_domicilio_direccion_calle: empty(datos.direccion_calle),
      datos_del_domicilio_direccion_numero: empty(datos.direccion_numero),
      datos_del_domicilio_direccion_colonia_o_barrio: empty(datos.direccion_colonia),
      datos_del_domicilio_direccion_codigo_postal: empty(datos.codigo_postal),

      // Actividad económica
      actividad_economica_ocupacion: empty(datos.ocupacion),
      actividad_economica_sector: empty(datos.sector),
      actividad_economica_negocio: empty(datos.negocio),
      actividad_economica_negocio_a_emprender: empty(datos.negocio_a_emprender),
      actividad_economica_ubicacion_del_negocio: empty(datos.ubicacion_negocio),
      actividad_economica_local: empty(datos.local),
      actividad_economica_anios_en_el_oficio: empty(datos.anios_oficio),
      actividad_economica_anios_en_el_negocio: empty(datos.anios_negocio),
      actividad_economica_numero_de_trabajadores: empty(datos.numero_trabajadores),
      actividad_economica_horario: empty(datos.horario),
      actividad_economica_telefono: empty(datos.telefono_trabajo),
      actividad_economica_pais: empty(datos.trabajo_pais),
      actividad_economica_estado: empty(datos.trabajo_estado),
      actividad_economica_municipio: empty(datos.trabajo_municipio),
      actividad_economica_localidad: empty(datos.trabajo_localidad),
      actividad_economica_referencia_de_localizacion: empty(datos.trabajo_referencia),
      actividad_economica_direccion_calle: empty(datos.trabajo_calle),
      actividad_economica_direccion_numero: empty(datos.trabajo_numero),
      actividad_economica_direccion_colonia_o_barrio: empty(datos.trabajo_colonia),
      actividad_economica_direccion_codigo_postal: empty(datos.trabajo_codigo_postal),
      actividad_economica_trabaja_lunes: boolX(datos.trabaja_lunes),
      actividad_economica_trabaja_martes: boolX(datos.trabaja_martes),
      actividad_economica_trabaja_miercoles: boolX(datos.trabaja_miercoles),
      actividad_economica_trabaja_jueves: boolX(datos.trabaja_jueves),
      actividad_economica_trabaja_viernes: boolX(datos.trabaja_viernes),
      actividad_economica_trabaja_sabado: boolX(datos.trabaja_sabado),
      actividad_economica_trabaja_domingo: boolX(datos.trabaja_domingo),
      actividad_economica_tiene_otro_ingreso: boolX(datos.tiene_otro_ingreso),
      actividad_economica_tiene_otro_ingreso_cual: empty(datos.otro_ingreso_cual),

      // Obligado solidario
      obligado_solidario_primer_nombre: empty(datos.obligado_primer_nombre),
      obligado_solidario_segundo_nombre: empty(datos.obligado_segundo_nombre),
      obligado_solidario_apellido_paterno: empty(datos.obligado_apellido_paterno),
      obligado_solidario_apellido_materno: empty(datos.obligado_apellido_materno),
      obligado_solidario_fecha_de_nacimiento: empty(datos.obligado_fecha_nacimiento),
      obligado_solidario_estado_de_nacimiento: empty(datos.obligado_estado_nacimiento),
      obligado_solidario_escolaridad: empty(datos.obligado_escolaridad),
      obligado_solidario_parentesco: empty(datos.obligado_parentesco),
      obligado_solidario_lugar_de_trabajo: empty(datos.obligado_lugar_trabajo),
      obligado_solidario_actividad_u_ocupacion: empty(datos.obligado_ocupacion),
      obligado_solidario_clave_de_elector: empty(datos.obligado_clave_elector),
      obligado_solidario_curp: empty(datos.obligado_curp),

      // Beneficiario
      datos_del_beneficiario_nombres: empty(datos.beneficiario_nombres),
      datos_del_beneficiario_apellidos: empty(datos.beneficiario_apellidos),
      datos_del_beneficiario_fecha_de_nacimiento: empty(datos.beneficiario_fecha_nacimiento),
      datos_del_beneficiario_sexo: empty(datos.beneficiario_sexo),
      datos_del_beneficiario_parentesco: empty(datos.beneficiario_parentesco),
      datos_del_beneficiario_participacion: empty(datos.beneficiario_participacion),
      datos_del_beneficiario_direccion: empty(datos.beneficiario_direccion),

      // Declaratorias
      declaratorias_actuo_a_nombre_y_cuenta_propia: boolX(datos.actuo_cuenta_propia),
      declaratorias_actuo_a_nombre_de_un_tercero: boolX(datos.actuo_nombre_tercero),
      declaratorias_tengo_relacion: boolX(datos.tengo_relacion),
      declaratorias_soy_accionista: boolX(datos.soy_accionista),
      declaratorias_autorizo_a_financiera_mandarme_informacion: boolX(datos.autorizo_informacion),
      declaratorias_como_se_entero_o_quien_recomendo: empty(datos.como_se_entero),
      declaratorias_tercero_primer_nombre: empty(datos.tercero_primer_nombre),
      declaratorias_tercero_segundo_nombre: empty(datos.tercero_segundo_nombre),
      declaratorias_tercero_apellido_paterno: empty(datos.tercero_apellido_paterno),
      declaratorias_tercero_apellido_materno: empty(datos.tercero_apellido_materno)
    };

    // Transformar true/false en X según el normalizador
    for (const [key, value] of Object.entries(mappedData)) {
      if (value === true || value === "true") {
        mappedData[key] = "X";
      } else if (typeof value === 'boolean' && value === false) {
        mappedData[key] = "";
      }
    }

    return {
      data: {
        type: 'excel',
        template: 'general',
        ...mappedData
      },
      template: 'general'
    };
  }

  /**
   * Mapeo para Visita Domiciliaria (Word)
   */
  mapVisitaDomiciliaria(datos) {
    const clean = (v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "string") {
        const s = v.trim();
        if (!s) return "";
        const low = s.toLowerCase();
        if (low === "null" || low === "undefined") return "";
        return s;
      }
      if (typeof v === "number") return String(v);
      if (typeof v === "boolean") return v;
      return String(v);
    };

    const boolToX = (v) => {
      if (v === true) return "X";
      if (v === false) return "";
      const s = clean(v).toLowerCase();
      if (!s) return "";
      if (["true", "1", "si", "sí", "x"].includes(s)) return "X";
      if (["false", "0", "no"].includes(s)) return "";
      return "";
    };

    const joinNonEmpty = (parts, sep = " ") => {
      return parts.map(clean).filter(Boolean).join(sep).trim();
    };

    const mappedData = {
      wa_id: clean(datos.wa_id),
      codigo_de_prospecto: clean(datos.codigo_de_prospecto) || clean(datos.id_expediente),

      nombre_del_cliente: joinNonEmpty([
        datos.primer_nombre,
        datos.segundo_nombre,
        datos.primer_apellido,
        datos.segundo_apellido
      ]),

      fecha: clean(datos.fecha_visita) || new Date().toLocaleDateString('es-MX'),
      grupo_al_que_pertenece: clean(datos.grupo),
      asesor: clean(datos.nombre_asesor),
      sucursal: clean(datos.sucursal_asesor),

      // Dirección
      direccion_vialidad: clean(datos.direccion_calle),
      direccion_numero: clean(datos.direccion_numero),
      direccion_colonia: clean(datos.direccion_colonia),
      direccion_ciudad: clean(datos.direccion_ciudad),
      direccion_municipio: clean(datos.direccion_municipio),
      direccion_estado: clean(datos.direccion_provincia),
      direccion_codigo_postal: clean(datos.codigo_postal),

      direccion_coincide_si: boolToX(datos.direccion_coincide_si),
      direccion_coincide_no: boolToX(datos.direccion_coincide_no),

      observaciones_domicilio_del_cliente: clean(datos.observaciones_domicilio),
      caracteristicas_principales_de_la_casa: clean(datos.la_casa_es),
      calles_entre_las_que_se_encuentra_el_domicilio: clean(datos.calles_entre_domicilio),
      lineas_o_rutas_de_transporte_para_llegar_a_domicilio: clean(datos.rutas_transporte),
      tiempo_aproximado_para_llegar_a_domicilio: clean(datos.tiempo_llegar),
      principales_referencias_de_ubicacion_del_domicilio: clean(datos.referencias_ubicacion),
      tiempo_de_vivir_en_domicilio: clean(datos.tiempo_vivir_domicilio),
      nombre_de_propietario_de_la_casa: clean(datos.propietario_casa),

      negocio_misma_direccion_si: boolToX(datos.negocio_misma_direccion_si),
      negocio_misma_direccion_no: boolToX(datos.negocio_misma_direccion_no),
      negocio_misma_direccion_no_direccion_completa: clean(datos.direccion_negocio_completa),

      ubicacion_domicilio: clean(datos.ubicacion_domicilio),
      ubicacion_negocio: clean(datos.ubicacion_negocio)
    };

    return {
      data: mappedData,
      template: 'Visita domiciliaria con etiquetas.docx'
    };
  }

  /**
   * Mapeo para Evaluación Económica Simple (Excel)
   */
  mapEvaluacionEconomicaSimple(datos) {
    const cleanVal = (v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "string") {
        const t = v.trim();
        if (t === "") return "";
        const low = t.toLowerCase();
        if (low === "null" || low === "undefined") return "";
        return t;
      }
      return v;
    };

    const mappedData = {
      // Campos excluidos
      sucursal: cleanVal(datos.sucursal_asesor),
      fecha: cleanVal(datos.fecha_evaluacion) || new Date().toLocaleDateString('es-MX'),
      nombre_del_cliente: [datos.primer_nombre, datos.segundo_nombre, datos.primer_apellido, datos.segundo_apellido].filter(Boolean).join(' '),
      secuencia: cleanVal(datos.secuencia),
      actividad_principal: cleanVal(datos.actividad_principal),
      grupo: cleanVal(datos.grupo),
      BC_Score: cleanVal(datos.bc_score),
      ICC: cleanVal(datos.icc),
      No_Hit: cleanVal(datos.no_hit),
      tipo_de_solicitante: cleanVal(datos.tipo_solicitante),
      monto_solicitado: cleanVal(datos.monto_solicitado),
      cuota_solicitada: cleanVal(datos.cuota_solicitada),

      // Etiquetas de ventas
      concepto_de_venta_1: cleanVal(datos.concepto_venta_1),
      concepto_de_venta_2: cleanVal(datos.concepto_venta_2),
      concepto_de_venta_3: cleanVal(datos.concepto_venta_3),
      concepto_de_venta_4: cleanVal(datos.concepto_venta_4),
      concepto_de_venta_5: cleanVal(datos.concepto_venta_5),
      concepto_de_venta_6: cleanVal(datos.concepto_venta_6),

      venta_1: cleanVal(datos.venta_1),
      venta_2: cleanVal(datos.venta_2),
      venta_3: cleanVal(datos.venta_3),
      venta_4: cleanVal(datos.venta_4),
      venta_5: cleanVal(datos.venta_5),
      venta_6: cleanVal(datos.venta_6),

      venta_semanal_1: cleanVal(datos.venta_semanal_1),
      venta_semanal_2: cleanVal(datos.venta_semanal_2),
      venta_semanal_3: cleanVal(datos.venta_semanal_3),
      venta_semanal_4: cleanVal(datos.venta_semanal_4),
      venta_semanal_5: cleanVal(datos.venta_semanal_5),
      venta_semanal_6: cleanVal(datos.venta_semanal_6),

      venta_quincenal_1: cleanVal(datos.venta_quincenal_1),
      venta_quincenal_2: cleanVal(datos.venta_quincenal_2),
      venta_quincenal_3: cleanVal(datos.venta_quincenal_3),
      venta_quincenal_4: cleanVal(datos.venta_quincenal_4),
      venta_quincenal_5: cleanVal(datos.venta_quincenal_5),
      venta_quincenal_6: cleanVal(datos.venta_quincenal_6),

      venta_mensual_1: cleanVal(datos.venta_mensual_1),
      venta_mensual_2: cleanVal(datos.venta_mensual_2),
      venta_mensual_3: cleanVal(datos.venta_mensual_3),
      venta_mensual_4: cleanVal(datos.venta_mensual_4),
      venta_mensual_5: cleanVal(datos.venta_mensual_5),
      venta_mensual_6: cleanVal(datos.venta_mensual_6),

      // Costos y gastos
      costo_1: cleanVal(datos.costo_1),
      costo_2: cleanVal(datos.costo_2),
      costo_3: cleanVal(datos.costo_3),
      costo_4: cleanVal(datos.costo_4),
      costo_5: cleanVal(datos.costo_5),
      costo_6: cleanVal(datos.costo_6),

      gastos_personales: cleanVal(datos.gastos_personales),
      gastos_generales: cleanVal(datos.gastos_generales),
      gastos_financieros: cleanVal(datos.gastos_financieros),
      otros_gastos: cleanVal(datos.otros_gastos),
      costo_de_ventas: cleanVal(datos.costo_ventas),
      utilidad_bruta: cleanVal(datos.utilidad_bruta),
      utilidad_neta: cleanVal(datos.utilidad_neta),

      // Ingresos de ganancia
      ingreso_de_ganancia_1: cleanVal(datos.ingreso_ganancia_1),
      ingreso_de_ganancia_2: cleanVal(datos.ingreso_ganancia_2),
      ingreso_de_ganancia_3: cleanVal(datos.ingreso_ganancia_3),
      ingreso_de_ganancia_4: cleanVal(datos.ingreso_ganancia_4),
      ingreso_de_ganancia_5: cleanVal(datos.ingreso_ganancia_5),
      ingreso_de_ganancia_6: cleanVal(datos.ingreso_ganancia_6),

      // Balance
      inventarios_activo: cleanVal(datos.inventarios_activo),
      caja_efectivo_activo: cleanVal(datos.caja_efectivo_activo),
      ahorro_bancos_activo: cleanVal(datos.ahorro_bancos_activo),
      cuentas_por_cobrar_activo: cleanVal(datos.cuentas_cobrar_activo),
      inventarios_pasivo: cleanVal(datos.inventarios_pasivo),
      mobiliario_maquinaria_equipo_activo: cleanVal(datos.mobiliario_activo),
      mobiliario_maquinaria_equipo_pasivo: cleanVal(datos.mobiliario_pasivo),
      local_u_otros_bienes_del_negocio_activo: cleanVal(datos.local_activo),
      local_u_otros_bienes_del_negocio_pasivo: cleanVal(datos.local_pasivo),

      comentarios_y_observaciones_adicionales: cleanVal(datos.comentarios_observaciones),
      monto_mayor_credito_obtenido: cleanVal(datos.monto_mayor_credito),
      monto_credito_anterior: cleanVal(datos.monto_credito_anterior),
      cuota_anterior: cleanVal(datos.cuota_anterior),
      pago_a_la_semana: cleanVal(datos.pago_semanal)
    };

    return {
      data: {
        type: "excel",
        formato: "evaluacion_economica",
        ...mappedData
      },
      template: 'evaluacion_economica'
    };
  }

  /**
   * Mapeo para Obligado Solidario (Word)
   */
  mapObligadoSolidario(datos) {
    const mappedData = {
      type: 'word',
      template: 'obligado_solidario',
      codigo: datos.codigo_de_prospecto || datos.id_expediente || '',
      fecha: datos.fecha || new Date().toLocaleDateString('es-MX'),
      wa_id: datos.wa_id || null,
      codigo_de_prospecto: datos.codigo_de_prospecto || datos.id_expediente || null,

      // Obligado
      'obligado.primer_nombre': datos.obligado_primer_nombre || '',
      'obligado.segundo_nombre': datos.obligado_segundo_nombre || '',
      'obligado.apellido_paterno': datos.obligado_apellido_paterno || '',
      'obligado.apellido_materno': datos.obligado_apellido_materno || '',
      'obligado.clave_de_elector': datos.obligado_clave_elector || '',
      'obligado.CURP': datos.obligado_curp || '',
      'obligado.RFC': datos.obligado_rfc || '',
      'obligado.firma_electronica': datos.obligado_firma_electronica || '',
      'obligado.nacionalidad': datos.obligado_nacionalidad || '',
      'obligado.pais_de_nacimiento': datos.obligado_pais_nacimiento || '',
      'obligado.estado_de_nacimiento': datos.obligado_estado_nacimiento || '',
      'obligado.fecha_de_nacimiento': datos.obligado_fecha_nacimiento || '',
      'obligado.estado_civil': datos.obligado_estado_civil || '',
      'obligado.depentientes_economicos': datos.obligado_dependientes_economicos || '',
      'obligado.sexo': datos.obligado_sexo || '',
      'obligado.escolaridad': datos.obligado_escolaridad || '',
      'obligado.actividad': datos.obligado_actividad || '',
      'obligado.profesion': datos.obligado_profesion || '',
      'obligado.ocupacion': datos.obligado_ocupacion || '',

      // Domicilio
      'domicilio.direccion_calle': datos.obligado_direccion_calle || '',
      'domicilio.direccion_numero': datos.obligado_direccion_numero || '',
      'domicilio.direccion_colonia': datos.obligado_direccion_colonia || '',
      'domicilio.direccion_ciudad': datos.obligado_direccion_ciudad || '',
      'domicilio.codigo_postal': datos.obligado_codigo_postal || '',
      'domicilio.municipio': datos.obligado_municipio || '',
      'domicilio.estado': datos.obligado_estado || '',
      'domicilio.pais': datos.obligado_pais || '',
      'domicilio.referecia_localizacion': datos.obligado_referencia_localizacion || '',
      'domicilio.la_casa_es': datos.obligado_la_casa_es || '',
      'domicilio.telefono': datos.obligado_telefono || '',

      // Cargo público
      'cargo_publico.si': datos.obligado_cargo_publico_si || '',
      'cargo_publico.no': datos.obligado_cargo_publico_no || '',
      'cargo_publico.familiares.si': datos.obligado_cargo_publico_familiares_si || '',
      'cargo_publico.familiares.no': datos.obligado_cargo_publico_familiares_no || '',

      // Protesta
      'protesta.es_accionista': datos.obligado_es_accionista || '',
      'protesta.tiene_relacion_con_accionista': datos.obligado_tiene_relacion_accionista || ''
    };

    return {
      data: mappedData,
      template: 'Fichadeidentificaciondelobligadosolidarioconetiquetas.docx'
    };
  }

  /**
   * Mapeo para Aval (Word)
   */
  mapAval(datos) {
    const mappedData = {
      type: 'word',
      template: 'ficha_aval',
      codigo: datos.codigo_de_prospecto || datos.id_expediente || '',
      fecha: datos.fecha || new Date().toLocaleDateString('es-MX'),
      wa_id: datos.wa_id || null,
      codigo_de_prospecto: datos.codigo_de_prospecto || datos.id_expediente || null,
      numero_aval: datos.numero_aval || 1,

      // Aval
      'aval.primer_nombre': datos.aval_primer_nombre || '',
      'aval.segundo_nombre': datos.aval_segundo_nombre || '',
      'aval.apellido_paterno': datos.aval_apellido_paterno || '',
      'aval.apellido_materno': datos.aval_apellido_materno || '',
      'aval.clave_de_elector': datos.aval_clave_elector || '',
      'aval.CURP': datos.aval_curp || '',
      'aval.RFC': datos.aval_rfc || '',
      'aval.firma_electronica': datos.aval_firma_electronica || '',
      'aval.nacionalidad': datos.aval_nacionalidad || '',
      'aval.pais_de_nacimiento': datos.aval_pais_nacimiento || '',
      'aval.estado_de_nacimiento': datos.aval_estado_nacimiento || '',
      'aval.fecha_de_nacimiento': datos.aval_fecha_nacimiento || '',
      'aval.estado_civil': datos.aval_estado_civil || '',
      'aval.depentientes_economicos': datos.aval_dependientes_economicos || '',
      'aval.sexo': datos.aval_sexo || '',
      'aval.escolaridad': datos.aval_escolaridad || '',
      'aval.actividad': datos.aval_actividad || '',
      'aval.profesion': datos.aval_profesion || '',
      'aval.ocupacion': datos.aval_ocupacion || '',

      // Domicilio
      'domicilio.direccion_calle': datos.aval_direccion_calle || '',
      'domicilio.direccion_numero': datos.aval_direccion_numero || '',
      'domicilio.direccion_colonia': datos.aval_direccion_colonia || '',
      'domicilio.direccion_ciudad': datos.aval_direccion_ciudad || '',
      'domicilio.codigo_postal': datos.aval_codigo_postal || '',
      'domicilio.municipio': datos.aval_municipio || '',
      'domicilio.estado': datos.aval_estado || '',
      'domicilio.pais': datos.aval_pais || '',
      'domicilio.referecia_localizacion': datos.aval_referencia_localizacion || '',
      'domicilio.la_casa_es': datos.aval_la_casa_es || '',
      'domicilio.telefono': datos.aval_telefono || '',

      // Pareja
      'pareja.apellido_paterno': datos.aval_pareja_apellido_paterno || '',
      'pareja.apellido_materno': datos.aval_pareja_apellido_materno || '',
      'pareja.primer_nombre': datos.aval_pareja_primer_nombre || '',
      'pareja.segundo_nombre': datos.aval_pareja_segundo_nombre || '',
      'pareja.estado_de_nacimiento': datos.aval_pareja_estado_nacimiento || '',
      'pareja.fecha_de_nacimiento': datos.aval_pareja_fecha_nacimiento || '',
      'pareja.ocupacion': datos.aval_pareja_ocupacion || '',
      'pareja.lugar_de_trabajo': datos.aval_pareja_lugar_trabajo || '',
      'pareja.clave_de_elector': datos.aval_pareja_clave_elector || '',
      'pareja.CURP': datos.aval_pareja_curp || '',
      'pareja.escolaridad': datos.aval_pareja_escolaridad || '',

      // Cargo público
      'cargo_publico.si': datos.aval_cargo_publico_si || '',
      'cargo_publico.no': datos.aval_cargo_publico_no || '',
      'cargo_publico.familiares.si': datos.aval_cargo_publico_familiares_si || '',
      'cargo_publico.familiares.no': datos.aval_cargo_publico_familiares_no || '',

      // Protesta
      'protesta.es_accionista': datos.aval_es_accionista || '',
      'protesta.tiene_relacion_con_accionista': datos.aval_tiene_relacion_accionista || ''
    };

    return {
      data: mappedData,
      template: 'ficha_de_identificacion_del_aval_con_etiquetas.docx'
    };
  }

  /**
   * Mapeo para Scoring Con Historial Crediticio (Excel)
   */
  mapScoringConHC(datos) {
    const cleanVal = (v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "string") {
        const t = v.trim();
        if (t === "") return "";
        const low = t.toLowerCase();
        if (low === "null" || low === "undefined") return "";
        return t;
      }
      return v;
    };

    // Mapeo basado en el excelService existente
    const mappedData = {
      // Datos básicos del cliente
      codigo_de_prospecto: cleanVal(datos.codigo_de_prospecto) || cleanVal(datos.id_expediente),
      nombre: [datos.primer_nombre, datos.segundo_nombre, datos.primer_apellido, datos.segundo_apellido].filter(Boolean).join(' '),
      apellido_paterno: cleanVal(datos.primer_apellido),
      apellido_materno: cleanVal(datos.segundo_apellido),
      telefono: cleanVal(datos.telefono),
      email: cleanVal(datos.correo),
      curp: cleanVal(datos.curp || datos.cedula),
      fecha_nacimiento: cleanVal(datos.fecha_nacimiento),
      edad: cleanVal(datos.edad),
      estado_civil: cleanVal(datos.estado_civil),
      sexo: cleanVal(datos.sexo),

      // Dirección
      calle: cleanVal(datos.direccion_calle),
      numero: cleanVal(datos.direccion_numero),
      colonia: cleanVal(datos.direccion_colonia),
      codigo_postal: cleanVal(datos.codigo_postal),
      municipio: cleanVal(datos.municipio),
      estado: cleanVal(datos.direccion_provincia),

      // Actividad económica
      ocupacion: cleanVal(datos.ocupacion),
      anos_en_el_negocio: cleanVal(datos.anios_negocio),
      la_casa_es: cleanVal(datos.la_casa_es),

      // Evaluación económica
      cuanto_ganas: cleanVal(datos.cuanto_ganas),
      cuanto_gastas: cleanVal(datos.cuanto_gastas),
      pagos_mensuales_creditos: cleanVal(datos.pagos_mensuales_creditos),
      egresos_mensuales: cleanVal(datos.egresos_mensuales),

      // Scoring específico (CON HC)
      calc_bcscore: cleanVal(datos.bc_score) || cleanVal(datos.calc_bcscore),
      'buro.BC_score': cleanVal(datos.bc_score),
      'buro.ICC': cleanVal(datos.icc),
      'buro.no_hit': cleanVal(datos.no_hit),

      // Elección final
      ultima_oferta: cleanVal(datos.monto_aceptado),
      monto_aceptado: cleanVal(datos.monto_aceptado),
      calc_capacidad_semanal: cleanVal(datos.pago_semanal),
      pago_semanal: cleanVal(datos.pago_semanal),

      // Referencias
      referencia1_nombre: cleanVal(datos.referencia1_nombre),
      referencia1_telefono: cleanVal(datos.referencia1_telefono),
      referencia2_nombre: cleanVal(datos.referencia2_nombre),
      referencia2_telefono: cleanVal(datos.referencia2_telefono),

      // Metadata
      wa_id: cleanVal(datos.wa_id)
    };

    return {
      data: mappedData,
      template: 'con_HC'
    };
  }

  /**
   * Mapeo para Scoring Sin Historial Crediticio (Excel)
   */
  mapScoringSinHC(datos) {
    const cleanVal = (v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "string") {
        const t = v.trim();
        if (t === "") return "";
        const low = t.toLowerCase();
        if (low === "null" || low === "undefined") return "";
        return t;
      }
      return v;
    };

    // Similar a con HC pero sin algunos campos específicos de historial
    const mappedData = {
      // Datos básicos del cliente
      codigo_de_prospecto: cleanVal(datos.codigo_de_prospecto) || cleanVal(datos.id_expediente),
      nombre: [datos.primer_nombre, datos.segundo_nombre, datos.primer_apellido, datos.segundo_apellido].filter(Boolean).join(' '),
      apellido_paterno: cleanVal(datos.primer_apellido),
      apellido_materno: cleanVal(datos.segundo_apellido),
      telefono: cleanVal(datos.telefono),
      email: cleanVal(datos.correo),
      curp: cleanVal(datos.curp || datos.cedula),
      fecha_nacimiento: cleanVal(datos.fecha_nacimiento),
      edad: cleanVal(datos.edad),
      estado_civil: cleanVal(datos.estado_civil),
      sexo: cleanVal(datos.sexo),

      // Dirección
      calle: cleanVal(datos.direccion_calle),
      numero: cleanVal(datos.direccion_numero),
      colonia: cleanVal(datos.direccion_colonia),
      codigo_postal: cleanVal(datos.codigo_postal),
      municipio: cleanVal(datos.municipio),
      estado: cleanVal(datos.direccion_provincia),

      // Actividad económica
      ocupacion: cleanVal(datos.ocupacion),
      anos_en_el_negocio: cleanVal(datos.anios_negocio),
      la_casa_es: cleanVal(datos.la_casa_es),

      // Evaluación económica
      cuanto_ganas: cleanVal(datos.cuanto_ganas),
      cuanto_gastas: cleanVal(datos.cuanto_gastas),
      pagos_mensuales_creditos: cleanVal(datos.pagos_mensuales_creditos),
      egresos_mensuales: cleanVal(datos.egresos_mensuales),

      // Scoring específico (SIN HC) - menos campos de buro
      calc_bcscore: cleanVal(datos.calc_bcscore) || cleanVal(datos.bc_score),

      // Elección final
      ultima_oferta: cleanVal(datos.monto_aceptado),
      monto_aceptado: cleanVal(datos.monto_aceptado),
      calc_capacidad_semanal: cleanVal(datos.pago_semanal),
      pago_semanal: cleanVal(datos.pago_semanal),

      // Referencias
      referencia1_nombre: cleanVal(datos.referencia1_nombre),
      referencia1_telefono: cleanVal(datos.referencia1_telefono),
      referencia2_nombre: cleanVal(datos.referencia2_nombre),
      referencia2_telefono: cleanVal(datos.referencia2_telefono),

      // Metadata
      wa_id: cleanVal(datos.wa_id)
    };

    return {
      data: mappedData,
      template: 'sin_HC'
    };
  }

  /**
   * Mapeo para Seguimiento de Crédito (Excel)
   */
  mapSeguimientoCredito(datos) {
    const cleanVal = (v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "string") {
        const t = v.trim();
        if (t === "") return "";
        const low = t.toLowerCase();
        if (low === "null" || low === "undefined") return "";
        return t;
      }
      return v;
    };

    const boolX = (v) => v === true || v === "true" || v === "X" ? "X" : "";

    // Mapeo según el template de seguimiento
    const mappedData = {
      // Básicos
      codigo_de_prospecto: cleanVal(datos.codigo_de_prospecto) || cleanVal(datos.id_expediente),
      wa_id: cleanVal(datos.wa_id),
      nombre_cliente: [datos.primer_nombre, datos.segundo_nombre, datos.primer_apellido, datos.segundo_apellido].filter(Boolean).join(' '),
      nombre_asesor: cleanVal(datos.nombre_asesor),

      // Fechas
      fecha_previo: cleanVal(datos.fecha_previo) || new Date().toLocaleDateString('es-MX'),
      fecha_post: cleanVal(datos.fecha_post),

      // Comentarios
      comentarios_previo: cleanVal(datos.comentarios_previo),
      comentarios_post: cleanVal(datos.comentarios_post),

      // Checkboxes de seguimiento (sí/no)
      monto_cliente_congruente_si: boolX(datos.monto_cliente_congruente_si),
      monto_cliente_congruente_no: boolX(datos.monto_cliente_congruente_no),
      riesgo_obligaciones_si: boolX(datos.riesgo_obligaciones_si),
      riesgo_obligaciones_no: boolX(datos.riesgo_obligaciones_no),
      riesgo_familiar_credito_si: boolX(datos.riesgo_familiar_credito_si),
      riesgo_familiar_credito_no: boolX(datos.riesgo_familiar_credito_no),
      enfermedad_riesgo_credito_si: boolX(datos.enfermedad_riesgo_credito_si),
      enfermedad_riesgo_credito_no: boolX(datos.enfermedad_riesgo_credito_no),
      autorizacion_gerente_si: boolX(datos.autorizacion_gerente_si),
      autorizacion_gerente_no: boolX(datos.autorizacion_gerente_no),
      problema_funcionamiento_si: boolX(datos.problema_funcionamiento_si),
      problema_funcionamiento_no: boolX(datos.problema_funcionamiento_no),
      mismo_aval_si: boolX(datos.mismo_aval_si),
      mismo_aval_no: boolX(datos.mismo_aval_no),
      credito_aplicado_si: boolX(datos.credito_aplicado_si),
      credito_aplicado_no: boolX(datos.credito_aplicado_no),
      negocio_cambios_si: boolX(datos.negocio_cambios_si),
      presenta_atrasos_si: boolX(datos.presenta_atrasos_si),
      presenta_atrasos_no: boolX(datos.presenta_atrasos_no),
      riesgo_recuperacion_si: boolX(datos.riesgo_recuperacion_si),
      riesgo_recuperacion_no: boolX(datos.riesgo_recuperacion_no),
      problema_cliente_si: boolX(datos.problema_cliente_si),
      problema_cliente_no: boolX(datos.problema_cliente_no),

      // Campos de inversión
      que_invertir_1: cleanVal(datos.que_invertir_1),
      que_invertir_2: cleanVal(datos.que_invertir_2),
      que_invertir_3: cleanVal(datos.que_invertir_3),
      que_invertir_4: cleanVal(datos.que_invertir_4),
      que_invertir_5: cleanVal(datos.que_invertir_5),
      valor_estimado_1: cleanVal(datos.valor_estimado_1),
      valor_estimado_2: cleanVal(datos.valor_estimado_2),
      valor_estimado_3: cleanVal(datos.valor_estimado_3),
      valor_estimado_4: cleanVal(datos.valor_estimado_4),
      valor_estimado_5: cleanVal(datos.valor_estimado_5)
    };

    return {
      data: mappedData,
      template: 'seguimiento'
    };
  }

  /**
   * Mapeo para Seguimiento Previo (Excel - alias de seguimiento_credito)
   */
  mapSeguimientoPrevio(datos) {
    // Usar el mismo mapeo que seguimiento_credito
    return this.mapSeguimientoCredito(datos);
  }

  /**
   * Construir nombre de archivo con formato: apellido_paterno_apellido_materno_primer_nombre_codigo_de_prospecto_ficha.(docx/xlsx)
   */
  buildCustomFileName(datos, tipoFicha, originalFileName) {
    try {
      console.log('[MULTIPLE-DOCS] 🔍 Datos para nombre personalizado:', {
        primer_apellido: datos.primer_apellido,
        segundo_apellido: datos.segundo_apellido,
        primer_nombre: datos.primer_nombre,
        codigo_de_prospecto: datos.codigo_de_prospecto,
        id_expediente: datos.id_expediente,
        tipoFicha: tipoFicha
      });

      // Extraer campos basándose en los mapeos
      const apellidoPaterno = this.cleanForFileName(datos.primer_apellido || '');
      const apellidoMaterno = this.cleanForFileName(datos.segundo_apellido || '');
      const primerNombre = this.cleanForFileName(datos.primer_nombre || '');
      const codigoProspecto = this.cleanForFileName(datos.codigo_de_prospecto || datos.id_expediente || '');

      console.log('[MULTIPLE-DOCS] 🔍 Campos limpios:', {
        apellidoPaterno, apellidoMaterno, primerNombre, codigoProspecto
      });

      // Determinar extensión basándose en el tipo de ficha
      let extension;
      switch (tipoFicha) {
        case 'visita_domiciliaria':
        case 'obligado_solidario':
        case 'aval':
          extension = 'docx';
          break;
        case 'identificacion_cliente':
        case 'evaluacion_economica_simple':
        case 'scoring_con_hc':
        case 'scoring_sin_hc':
        case 'seguimiento_credito':
        case 'seguimiento_previo':
        default:
          extension = 'xlsx';
          break;
      }

      // Construir nombre: apellido_paterno_apellido_materno_primer_nombre_codigo_de_prospecto_ficha.ext
      const partes = [apellidoPaterno, apellidoMaterno, primerNombre, codigoProspecto, tipoFicha].filter(Boolean);
      const nombreCustom = partes.join('_') + '.' + extension;

      console.log('[MULTIPLE-DOCS] 📄 Nombre personalizado generado:', nombreCustom);
      console.log('[MULTIPLE-DOCS] 📄 Nombre original era:', originalFileName);

      return nombreCustom || originalFileName;

    } catch (error) {
      console.error('[MULTIPLE-DOCS] ❌ Error creando nombre personalizado:', error.message);
      return originalFileName;
    }
  }

  /**
   * Limpiar string para nombre de archivo
   */
  cleanForFileName(str) {
    if (!str) return '';
    return String(str)
      .trim()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '') // Remover acentos
      .replace(/[^a-zA-Z0-9]/g, '_') // Reemplazar caracteres especiales con _
      .replace(/_{2,}/g, '_') // Reemplazar múltiples _ con uno solo
      .replace(/^_+|_+$/g, '') // Remover _ al inicio y final
      .toUpperCase();
  }
}

module.exports = new MultipleDocumentsService();