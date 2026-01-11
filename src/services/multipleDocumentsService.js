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
      scoring: this.mapScoringAuto.bind(this),
      scoring_con_hc: this.mapScoringConHC.bind(this),
      scoring_sin_hc: this.mapScoringSinHC.bind(this),
      scoring_con_etiquetas: this.mapScoringConEtiquetas.bind(this),
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
        case 'scoring':
        case 'scoring_con_hc':
        case 'scoring_sin_hc':
        case 'scoring_con_etiquetas':
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
    const empty = (v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "string") {
        const str = v.trim();
        if (str === "" || str.toLowerCase() === "undefined" || str.toLowerCase() === "null") return "";
      }
      return v;
    };
    const boolX = (v) => v === true || v === "true" || v === "X" ? "X" : "";

    const nombreCliente = cleanVal(datos.nombre_cliente) || [
      datos.primer_nombre,
      datos.segundo_nombre,
      datos.primer_apellido || datos.apellido_paterno,
      datos.segundo_apellido || datos.apellido_materno
    ].filter(Boolean).join(' ');

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
        if (!s || s.toLowerCase() === "null" || s.toLowerCase() === "undefined") return "";
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

    const firstNonEmpty = (...values) => {
      for (const value of values) {
        const cleaned = clean(value);
        if (cleaned !== "") return cleaned;
      }
      return "";
    };

    const mappedData = {
      wa_id: clean(datos.wa_id),
      codigo_de_prospecto: clean(datos.codigo_de_prospecto) || clean(datos.id_expediente),

      nombre_del_cliente: firstNonEmpty(
        datos.nombre_cliente,
        joinNonEmpty([
          datos.primer_nombre || datos.cliente_primer_nombre,
          datos.segundo_nombre || datos.cliente_segundo_nombre,
          datos.primer_apellido || datos.apellido_paterno || datos.cliente_apellido_paterno,
          datos.segundo_apellido || datos.apellido_materno || datos.cliente_apellido_materno
        ])
      ),

      fecha: clean(datos.fecha_visita) || clean(datos.fecha) || new Date().toLocaleDateString('es-MX'),
      grupo_al_que_pertenece: clean(datos.grupo),
      asesor: clean(datos.nombre_asesor),
      sucursal: firstNonEmpty(datos.sucursal_asesor, datos.nombre_sucursal),

      // Dirección
      direccion_vialidad: firstNonEmpty(datos.direccion_calle, datos.datos_del_domicilio_direccion_calle),
      direccion_numero: firstNonEmpty(datos.direccion_numero, datos.datos_del_domicilio_direccion_numero),
      direccion_colonia: firstNonEmpty(datos.direccion_colonia, datos.datos_del_domicilio_direccion_colonia_o_barrio),
      direccion_ciudad: firstNonEmpty(datos.direccion_ciudad, datos.datos_del_domicilio_localidad),
      direccion_municipio: firstNonEmpty(datos.direccion_municipio, datos.datos_del_domicilio_municipio),
      direccion_estado: firstNonEmpty(datos.direccion_provincia, datos.datos_del_domicilio_estado),
      direccion_codigo_postal: firstNonEmpty(datos.codigo_postal, datos.datos_del_domicilio_direccion_codigo_postal),

      direccion_coincide_si: boolToX(datos.direccion_coincide_si),
      direccion_coincide_no: boolToX(datos.direccion_coincide_no),

      observaciones_domicilio_del_cliente: firstNonEmpty(
        datos.observaciones_domicilio,
        datos.observaciones_domicilio_del_cliente
      ),
      caracteristicas_principales_de_la_casa: firstNonEmpty(datos.la_casa_es, datos.datos_del_domicilio_la_casa_es),
      calles_entre_las_que_se_encuentra_el_domicilio: firstNonEmpty(
        datos.calles_entre_domicilio,
        datos.calles_entre_las_que_se_encuentra_el_domicilio
      ),
      lineas_o_rutas_de_transporte_para_llegar_a_domicilio: firstNonEmpty(
        datos.rutas_transporte,
        datos.lineas_o_rutas_de_transporte_para_llegar_a_domicilio
      ),
      tiempo_aproximado_para_llegar_a_domicilio: firstNonEmpty(
        datos.tiempo_llegar,
        datos.tiempo_aproximado_para_llegar_a_domicilio
      ),
      principales_referencias_de_ubicacion_del_domicilio: firstNonEmpty(
        datos.referencias_ubicacion,
        datos.datos_del_domicilio_referencia_de_localizacion
      ),
      tiempo_de_vivir_en_domicilio: firstNonEmpty(
        datos.tiempo_vivir_domicilio,
        datos.tiempo_de_vivir_en_domicilio
      ),
      nombre_de_propietario_de_la_casa: firstNonEmpty(
        datos.propietario_casa,
        datos.nombre_de_propietario_de_la_casa
      ),

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
        if (t === "" || t.toLowerCase() === "null" || t.toLowerCase() === "undefined") return "";
        return t;
      }
      return v;
    };

    const pickVal = (...values) => {
      for (const value of values) {
        const cleaned = cleanVal(value);
        if (cleaned !== "") return cleaned;
      }
      return "";
    };

    const nombreCliente = pickVal(
      datos.nombre_cliente,
      [datos.primer_nombre, datos.segundo_nombre, datos.primer_apellido, datos.segundo_apellido].filter(Boolean).join(' '),
      [datos.primer_nombre, datos.segundo_nombre, datos.apellido_paterno, datos.apellido_materno].filter(Boolean).join(' '),
      [datos.cliente_primer_nombre, datos.cliente_segundo_nombre, datos.cliente_apellido_paterno, datos.cliente_apellido_materno].filter(Boolean).join(' ')
    );

    const mappedData = {
      // Campos excluidos
      sucursal: pickVal(datos.sucursal_asesor, datos.nombre_sucursal),
      fecha: pickVal(datos.fecha_evaluacion, datos.fecha, new Date().toLocaleDateString('es-MX')),
      nombre_del_cliente: nombreCliente,
      secuencia: cleanVal(datos.secuencia),
      actividad_principal: pickVal(datos.actividad_principal, datos.actividad_economica_ocupacion),
      grupo: pickVal(datos.grupo, datos.calc_grupo),
      BC_Score: pickVal(datos.bc_score, datos.calc_bcscore),
      ICC: cleanVal(datos.icc),
      No_Hit: cleanVal(datos.no_hit),
      tipo_de_solicitante: cleanVal(datos.tipo_solicitante),
      monto_solicitado: cleanVal(datos.monto_solicitado),
      cuota_solicitada: cleanVal(datos.cuota_solicitada),

      // Etiquetas de ventas
      concepto_de_venta_1: pickVal(datos.concepto_de_venta_1, datos.concepto_venta_1),
      concepto_de_venta_2: pickVal(datos.concepto_de_venta_2, datos.concepto_venta_2),
      concepto_de_venta_3: pickVal(datos.concepto_de_venta_3, datos.concepto_venta_3),
      concepto_de_venta_4: pickVal(datos.concepto_de_venta_4, datos.concepto_venta_4),
      concepto_de_venta_5: pickVal(datos.concepto_de_venta_5, datos.concepto_venta_5),
      concepto_de_venta_6: pickVal(datos.concepto_de_venta_6, datos.concepto_venta_6),

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
      costo_de_ventas: pickVal(datos.costo_de_ventas, datos.costo_ventas),
      utilidad_bruta: cleanVal(datos.utilidad_bruta),
      utilidad_neta: cleanVal(datos.utilidad_neta),

      // Ingresos de ganancia / Porcentajes de ganancia
      ingreso_de_ganancia_1: pickVal(datos.ingreso_de_ganancia_1, datos.ingreso_ganancia_1, datos.porcentaje_de_ganancia_1),
      ingreso_de_ganancia_2: pickVal(datos.ingreso_de_ganancia_2, datos.ingreso_ganancia_2, datos.porcentaje_de_ganancia_2),
      ingreso_de_ganancia_3: pickVal(datos.ingreso_de_ganancia_3, datos.ingreso_ganancia_3, datos.porcentaje_de_ganancia_3),
      ingreso_de_ganancia_4: pickVal(datos.ingreso_de_ganancia_4, datos.ingreso_ganancia_4, datos.porcentaje_de_ganancia_4),
      ingreso_de_ganancia_5: pickVal(datos.ingreso_de_ganancia_5, datos.ingreso_ganancia_5, datos.porcentaje_de_ganancia_5),
      ingreso_de_ganancia_6: pickVal(datos.ingreso_de_ganancia_6, datos.ingreso_ganancia_6, datos.porcentaje_de_ganancia_6),
      porcentaje_de_ganancia_1: pickVal(datos.porcentaje_de_ganancia_1, datos.ingreso_de_ganancia_1),
      porcentaje_de_ganancia_2: pickVal(datos.porcentaje_de_ganancia_2, datos.ingreso_de_ganancia_2),
      porcentaje_de_ganancia_3: pickVal(datos.porcentaje_de_ganancia_3, datos.ingreso_de_ganancia_3),
      porcentaje_de_ganancia_4: pickVal(datos.porcentaje_de_ganancia_4, datos.ingreso_de_ganancia_4),
      porcentaje_de_ganancia_5: pickVal(datos.porcentaje_de_ganancia_5, datos.ingreso_de_ganancia_5),
      porcentaje_de_ganancia_6: pickVal(datos.porcentaje_de_ganancia_6, datos.ingreso_de_ganancia_6),

      // Balance
      inventarios_activo: cleanVal(datos.inventarios_activo),
      caja_efectivo_activo: cleanVal(datos.caja_efectivo_activo),
      ahorro_bancos_activo: cleanVal(datos.ahorro_bancos_activo),
      cuentas_por_cobrar_activo: pickVal(datos.cuentas_por_cobrar_activo, datos.cuentas_cobrar_activo),
      inventarios_pasivo: cleanVal(datos.inventarios_pasivo),
      mobiliario_maquinaria_equipo_activo: pickVal(datos.mobiliario_maquinaria_equipo_activo, datos.mobiliario_activo),
      mobiliario_maquinaria_equipo_pasivo: pickVal(datos.mobiliario_maquinaria_equipo_pasivo, datos.mobiliario_pasivo),
      local_u_otros_bienes_del_negocio_activo: pickVal(datos.local_u_otros_bienes_del_negocio_activo, datos.local_activo),
      local_u_otros_bienes_del_negocio_pasivo: pickVal(datos.local_u_otros_bienes_del_negocio_pasivo, datos.local_pasivo),

      comentarios_y_observaciones_adicionales: pickVal(
        datos.comentarios_y_observaciones_adicionales,
        datos.comentarios_observaciones
      ),
      monto_mayor_credito_obtenido: pickVal(datos.monto_mayor_credito_obtenido, datos.monto_mayor_credito),
      monto_credito_anterior: cleanVal(datos.monto_credito_anterior),
      cuota_anterior: cleanVal(datos.cuota_anterior),
      pago_a_la_semana: pickVal(datos.pago_a_la_semana, datos.pago_semanal)
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
   * Mapeo para Scoring (auto: con HC o sin HC)
   */
  mapScoringAuto(datos) {
    const hasValue = (v) => v !== null && v !== undefined && !(typeof v === "string" && v.trim() === "");
    const tieneBuro = hasValue(datos.bc_score) || hasValue(datos.icc) || hasValue(datos.no_hit);
    return tieneBuro ? this.mapScoringConHC(datos) : this.mapScoringSinHC(datos);
  }

  /**
   * Mapeo para Scoring Con Historial Crediticio (Excel)
   */
  mapScoringConHC(datos) {
    const cleanVal = (v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "string") {
        const t = v.trim();
        if (t === "" || t.toLowerCase() === "null" || t.toLowerCase() === "undefined") return "";
        return t;
      }
      return v;
    };

    const pickVal = (...values) => {
      for (const value of values) {
        const cleaned = cleanVal(value);
        if (cleaned !== "") return cleaned;
      }
      return "";
    };

    const nombreCompleto = pickVal(
      datos.nombre_cliente,
      [datos.primer_nombre, datos.segundo_nombre, datos.primer_apellido, datos.segundo_apellido].filter(Boolean).join(' '),
      [datos.primer_nombre, datos.segundo_nombre, datos.apellido_paterno, datos.apellido_materno].filter(Boolean).join(' '),
      [datos.cliente_primer_nombre, datos.cliente_segundo_nombre, datos.cliente_apellido_paterno, datos.cliente_apellido_materno].filter(Boolean).join(' ')
    );

    // Mapeo basado en el excelService existente
    const mappedData = {
      // Datos básicos del cliente
      codigo_de_prospecto: cleanVal(datos.codigo_de_prospecto) || cleanVal(datos.id_expediente),
      nombre: nombreCompleto,
      apellido_paterno: pickVal(datos.primer_apellido, datos.apellido_paterno, datos.cliente_apellido_paterno),
      apellido_materno: pickVal(datos.segundo_apellido, datos.apellido_materno, datos.cliente_apellido_materno),
      telefono: pickVal(datos.telefono, datos.datos_del_domicilio_telefono),
      email: pickVal(datos.correo, datos.cliente_correo_electronico),
      curp: pickVal(datos.curp, datos.cliente_curp, datos.cedula),
      fecha_nacimiento: pickVal(datos.fecha_nacimiento, datos.fecha_de_nacimiento, datos.cliente_fecha_de_nacimiento),
      edad: cleanVal(datos.edad),
      estado_civil: cleanVal(datos.estado_civil),
      sexo: pickVal(datos.sexo, datos.cliente_sexo),

      // Dirección
      calle: pickVal(datos.direccion_calle, datos.datos_del_domicilio_direccion_calle),
      numero: pickVal(datos.direccion_numero, datos.datos_del_domicilio_direccion_numero),
      colonia: pickVal(datos.direccion_colonia, datos.datos_del_domicilio_direccion_colonia_o_barrio),
      codigo_postal: pickVal(
        datos.codigo_postal,
        datos.datos_del_domicilio_direccion_codigo_postal,
        datos.codigo_postal_cliente
      ),
      municipio: pickVal(datos.municipio, datos.datos_del_domicilio_municipio),
      estado: pickVal(datos.direccion_provincia, datos.datos_del_domicilio_estado, datos.estado_cliente),

      // Actividad económica
      ocupacion: pickVal(datos.ocupacion, datos.actividad_economica_ocupacion),
      anos_en_el_negocio: pickVal(datos.anios_negocio, datos.actividad_economica_anios_en_el_negocio),
      la_casa_es: pickVal(datos.la_casa_es, datos.datos_del_domicilio_la_casa_es),

      // Evaluación económica
      cuanto_ganas: cleanVal(datos.cuanto_ganas),
      cuanto_gastas: cleanVal(datos.cuanto_gastas),
      pagos_mensuales_creditos: cleanVal(datos.pagos_mensuales_creditos),
      egresos_mensuales: cleanVal(datos.egresos_mensuales),

      // Scoring específico (CON HC)
      calc_bcscore: pickVal(datos.bc_score, datos.calc_bcscore),
      'buro.BC_score': cleanVal(datos.bc_score),
      'buro.ICC': cleanVal(datos.icc),
      'buro.no_hit': cleanVal(datos.no_hit),

      // Elección final
      ultima_oferta: cleanVal(datos.monto_aceptado),
      monto_aceptado: cleanVal(datos.monto_aceptado),
      calc_capacidad_semanal: pickVal(datos.pago_semanal, datos.pago_a_la_semana),
      pago_semanal: pickVal(datos.pago_semanal, datos.pago_a_la_semana),

      // Referencias
      referencia1_nombre: pickVal(
        datos.referencia1_nombre,
        datos.primera_referencia_personal_nombre_completo
      ),
      referencia1_telefono: pickVal(
        datos.referencia1_telefono,
        datos.primera_referencia_personal_telefono
      ),
      referencia2_nombre: pickVal(
        datos.referencia2_nombre,
        datos.segunda_referencia_personal_nombre_completo
      ),
      referencia2_telefono: pickVal(
        datos.referencia2_telefono,
        datos.segunda_referencia_personal_telefono
      ),

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
        if (t === "" || t.toLowerCase() === "null" || t.toLowerCase() === "undefined") return "";
        return t;
      }
      return v;
    };

    const pickVal = (...values) => {
      for (const value of values) {
        const cleaned = cleanVal(value);
        if (cleaned !== "") return cleaned;
      }
      return "";
    };

    const nombreCompleto = pickVal(
      datos.nombre_cliente,
      [datos.primer_nombre, datos.segundo_nombre, datos.primer_apellido, datos.segundo_apellido].filter(Boolean).join(' '),
      [datos.primer_nombre, datos.segundo_nombre, datos.apellido_paterno, datos.apellido_materno].filter(Boolean).join(' '),
      [datos.cliente_primer_nombre, datos.cliente_segundo_nombre, datos.cliente_apellido_paterno, datos.cliente_apellido_materno].filter(Boolean).join(' ')
    );

    // Similar a con HC pero sin algunos campos específicos de historial
    const mappedData = {
      // Datos básicos del cliente
      codigo_de_prospecto: cleanVal(datos.codigo_de_prospecto) || cleanVal(datos.id_expediente),
      nombre: nombreCompleto,
      apellido_paterno: pickVal(datos.primer_apellido, datos.apellido_paterno, datos.cliente_apellido_paterno),
      apellido_materno: pickVal(datos.segundo_apellido, datos.apellido_materno, datos.cliente_apellido_materno),
      telefono: pickVal(datos.telefono, datos.datos_del_domicilio_telefono),
      email: pickVal(datos.correo, datos.cliente_correo_electronico),
      curp: pickVal(datos.curp, datos.cliente_curp, datos.cedula),
      fecha_nacimiento: pickVal(datos.fecha_nacimiento, datos.fecha_de_nacimiento, datos.cliente_fecha_de_nacimiento),
      edad: cleanVal(datos.edad),
      estado_civil: cleanVal(datos.estado_civil),
      sexo: pickVal(datos.sexo, datos.cliente_sexo),

      // Dirección
      calle: pickVal(datos.direccion_calle, datos.datos_del_domicilio_direccion_calle),
      numero: pickVal(datos.direccion_numero, datos.datos_del_domicilio_direccion_numero),
      colonia: pickVal(datos.direccion_colonia, datos.datos_del_domicilio_direccion_colonia_o_barrio),
      codigo_postal: pickVal(
        datos.codigo_postal,
        datos.datos_del_domicilio_direccion_codigo_postal,
        datos.codigo_postal_cliente
      ),
      municipio: pickVal(datos.municipio, datos.datos_del_domicilio_municipio),
      estado: pickVal(datos.direccion_provincia, datos.datos_del_domicilio_estado, datos.estado_cliente),

      // Actividad económica
      ocupacion: pickVal(datos.ocupacion, datos.actividad_economica_ocupacion),
      anos_en_el_negocio: pickVal(datos.anios_negocio, datos.actividad_economica_anios_en_el_negocio),
      la_casa_es: pickVal(datos.la_casa_es, datos.datos_del_domicilio_la_casa_es),

      // Evaluación económica
      cuanto_ganas: cleanVal(datos.cuanto_ganas),
      cuanto_gastas: cleanVal(datos.cuanto_gastas),
      pagos_mensuales_creditos: cleanVal(datos.pagos_mensuales_creditos),
      egresos_mensuales: cleanVal(datos.egresos_mensuales),

      // Scoring específico (SIN HC) - menos campos de buro
      calc_bcscore: pickVal(datos.calc_bcscore, datos.bc_score),

      // Elección final
      ultima_oferta: cleanVal(datos.monto_aceptado),
      monto_aceptado: cleanVal(datos.monto_aceptado),
      calc_capacidad_semanal: pickVal(datos.pago_semanal, datos.pago_a_la_semana),
      pago_semanal: pickVal(datos.pago_semanal, datos.pago_a_la_semana),

      // Referencias
      referencia1_nombre: pickVal(
        datos.referencia1_nombre,
        datos.primera_referencia_personal_nombre_completo
      ),
      referencia1_telefono: pickVal(
        datos.referencia1_telefono,
        datos.primera_referencia_personal_telefono
      ),
      referencia2_nombre: pickVal(
        datos.referencia2_nombre,
        datos.segunda_referencia_personal_nombre_completo
      ),
      referencia2_telefono: pickVal(
        datos.referencia2_telefono,
        datos.segunda_referencia_personal_telefono
      ),

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
        if (t === "" || t.toLowerCase() === "null" || t.toLowerCase() === "undefined") return "";
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
      nombre_cliente: nombreCliente,
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
   * Mapeo para Scoring Con Etiquetas (Excel)
   */
  mapScoringConEtiquetas(datos) {
    const cleanVal = (v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "string") {
        const t = v.trim();
        if (t === "" || t.toLowerCase() === "null" || t.toLowerCase() === "undefined") return "";
        return t;
      }
      return v;
    };

    const boolX = (v) => v === true || v === "true" || v === "X" ? "X" : "";

    // Mapeo basado en el mapfield_scoring.csv
    const mappedData = {
      // Datos básicos
      nombre_sucursal: cleanVal(datos.sucursal_asesor) || cleanVal(datos.sucursal),
      fecha: cleanVal(datos.fecha_scoring) || new Date().toLocaleDateString('es-MX'),
      nombre_cliente: [datos.primer_nombre, datos.segundo_nombre, datos.primer_apellido, datos.segundo_apellido].filter(Boolean).join(' '),
      secuencia_de_credito: cleanVal(datos.secuencia) || cleanVal(datos.secuencia_de_credito),

      // Tipo de vivienda
      'tipo_de_vivienda.propia': boolX(datos.vivienda_propia),
      'tipo_de_vivienda.rentada': boolX(datos.vivienda_rentada),
      'tipo_de_vivienda.habita_en_casa_de_familiar': boolX(datos.vivienda_familiar),
      'tipo_de_vivienda.prestada_y_compartida': boolX(datos.vivienda_prestada),
      'tipo_de_vivienda.rentada_y_compartida': boolX(datos.vivienda_rentada_compartida),

      // Tiempo de vivir en domicilio
      'tiempo_de_vivir_en_domicilio.mas_de_7_años': boolX(datos.tiempo_domicilio_mas_7),
      'tiempo_de_vivir_en_domicilio.entre_5_y_7_años': boolX(datos.tiempo_domicilio_5_7),
      'tiempo_de_vivir_en_domicilio.entre_3_y_5_años': boolX(datos.tiempo_domicilio_3_5),
      'tiempo_de_vivir_en_domicilio.entre_1_y_3_años': boolX(datos.tiempo_domicilio_1_3),
      'tiempo_de_vivir_en_domicilio.1_año_o_menos': boolX(datos.tiempo_domicilio_menos_1),

      // Impresión de situación o vivienda
      'impresion_de_situacion_o_vivienda.casa_con_ladrillo': boolX(datos.casa_ladrillo),
      'impresion_de_situacion_o_vivienda.casa_en_obra_gris': boolX(datos.casa_obra_gris),
      'impresion_de_situacion_o_vivienda.casa_en_obra_negra': boolX(datos.casa_obra_negra),
      'impresion_de_situacion_o_vivienda.casa_en_mal_estado': boolX(datos.casa_mal_estado),
      'impresion_de_situacion_o_vivienda.casa_en_mal_estado_y_condiciones_deficientes': boolX(datos.casa_condiciones_deficientes),

      // Edad del solicitante
      'edad_del_solicitante.entre_36_y_50_años': boolX(datos.edad_36_50),
      'edad_del_solicitante.entre_51_y_74_años': boolX(datos.edad_51_74),
      'edad_del_solicitante.entre_26_y_35_años': boolX(datos.edad_26_35),
      'edad_del_solicitante.entre_22_y_25_años': boolX(datos.edad_22_25),
      'edad_del_solicitante.entre_18_y_21_años': boolX(datos.edad_18_21),

      // Estado civil
      'estado_civil.casado_con_mas_de_3_dependientes': boolX(datos.casado_mas_3_dep),
      'estado_civil.casado_con_menos_de_3_dependientes': boolX(datos.casado_menos_3_dep),
      'estado_civil.union_libre_con_mas_de_3_años_juntos': boolX(datos.union_libre_mas_3),
      'estado_civil.union_libre_con_menos_de_3_años_juntos': boolX(datos.union_libre_menos_3),
      'estado_civil.separado_viudo_soltero_sin_dependientes': boolX(datos.separado_viudo_soltero),

      // Solicitante
      'solicitante.recomendado_por_mas_de_3_personas': boolX(datos.recomendado_mas_3),
      'solicitante.es_conocido_pero_no_personalmente': boolX(datos.conocido_no_personal),
      'solicitante.con_solo_2_referencias': boolX(datos.solo_2_referencias),
      'solicitante.con_dificultad_para_referencias_e_informacion_imprecisa': boolX(datos.referencias_imprecisas),
      'solicitante.con_dificultad_para_referencias_dudosas_y_comprometidas': boolX(datos.referencias_dudosas),

      // Tiempo negocio
      'tiempo_negocio.mas_de_5_años_con_mismo_giro': boolX(datos.negocio_mas_5_años),
      'tiempo_negocio.de_3_a_5_años_con_mismo_giro': boolX(datos.negocio_3_5_años),
      'tiempo_negocio.de_1_a_3_años_con_mismo_giro_o_similar': boolX(datos.negocio_1_3_años),
      'tiempo_negocio.menos_de_1_año': boolX(datos.negocio_menos_1_año),
      'tiempo_negocio.viene_de_otro_giro': boolX(datos.negocio_otro_giro),

      // Ubicación y tipo
      'ubicacion_y_tipo.negocio_fijo_local_propio': boolX(datos.negocio_local_propio),
      'ubicacion_y_tipo.negocio_fijo_local_rentado': boolX(datos.negocio_local_rentado),
      'ubicacion_y_tipo.negocio_semifijo': boolX(datos.negocio_semifijo),
      'ubicacion_y_tipo.negocio_ambulante': boolX(datos.negocio_ambulante),
      'ubicacion_y_tipo.venta_de_catalogo': boolX(datos.venta_catalogo),

      // Tipo de actividad
      'tipo_de_actividad.produccion_y_transformacion': boolX(datos.actividad_produccion),
      'tipo_de_actividad.comercio_y_servicios': boolX(datos.actividad_comercio),
      'tipo_de_actividad.artesanales_y_agropecuarias': boolX(datos.actividad_artesanal),
      'tipo_de_actividad.venta_por_catalogo': boolX(datos.actividad_catalogo),
      'tipo_de_actividad.transportista': boolX(datos.actividad_transporte),

      // Información financiera
      'informacion_financiera.entrega_estados_financieros': boolX(datos.entrega_estados_financieros),
      'informacion_financiera.muestra_facturas_que_acreditan_ingresos': boolX(datos.facturas_acreditan_ingresos),
      'informacion_financiera.muestra_facturas_que_no_acreditan_ingresos': boolX(datos.facturas_no_acreditan_ingresos),
      'informacion_financiera.sin_comprobantes_informacion_no_consistente': boolX(datos.sin_comprobantes_inconsistente),
      'informacion_financiera.respestas_evasivas_datos_sin_soporte': boolX(datos.respuestas_evasivas),

      // Historial crediticio interno
      'historial_crediticio.interno.0_atrasos_en_ultimo_credito': boolX(datos.historial_0_atrasos),
      'historial_crediticio.interno.1_a_5_dias_de_atraso_en_su_ultimo_credito': boolX(datos.historial_1_5_dias),
      'historial_crediticio.interno.6_a_15_dias_de_atraso_en_su_ultimo_credito': boolX(datos.historial_6_15_dias),
      'historial_crediticio.interno.mora_recurrente_o_mas_de_15_dias_de_atraso_en_su_ultimo_credito': boolX(datos.historial_mora_recurrente),
      'historial_crediticio.interno.cliente_nuevo': boolX(datos.historial_cliente_nuevo),

      // Historial crediticio externo
      'historial_crediticio.externo.BC_Score_igual_o_mayor_a_601.no_hit_mayor_a_650': boolX(datos.bc_score_601_mas),
      'historial_crediticio.externo.BC_Score_de_501_a_600.no_hit_de_601_a_650': boolX(datos.bc_score_501_600),
      'historial_crediticio.externo.BC_Score_de_401_a_500.no_hit_de_581_a_600': boolX(datos.bc_score_401_500),
      'historial_crediticio.externo.BC_Score_de_301_a_400.no_hit_de_561_a_580': boolX(datos.bc_score_301_400),
      'historial_crediticio.externo.BC_Score_menor_a_300.no_hit_igual_o_menor_a_560': boolX(datos.bc_score_menor_300),

      // Capacidad de pago
      'capacidad_de_pago.3_a_1_en_adelante': boolX(datos.capacidad_pago_3_1),
      'capacidad_de_pago.entre_2.5_a_1_y_2.9_a_1': boolX(datos.capacidad_pago_25_29),
      'capacidad_de_pago.entre_2_a_1_y_2.4_a_1': boolX(datos.capacidad_pago_2_24),
      'capacidad_de_pago.entre_1.5_a_1_y_1.9_a_1': boolX(datos.capacidad_pago_15_19),
      'capacidad_de_pago.igual_a_1.4_a_1': boolX(datos.capacidad_pago_14)
    };

    return {
      data: mappedData,
      template: 'scoring_etiquetas'
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
      const apellidoPaterno = this.cleanForFileName(
        datos.primer_apellido || datos.apellido_paterno || datos.cliente_apellido_paterno || ''
      );
      const apellidoMaterno = this.cleanForFileName(
        datos.segundo_apellido || datos.apellido_materno || datos.cliente_apellido_materno || ''
      );
      const primerNombre = this.cleanForFileName(
        datos.primer_nombre || datos.cliente_primer_nombre || ''
      );
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
        case 'scoring_con_etiquetas':
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
