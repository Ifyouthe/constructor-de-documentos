// =============================================
// ALMACENAMIENTO DE PLANTILLAS - CONSTRUCTOR DOCUMENTOS SUMATE
// Sustituye a src/config/supabase.js: las plantillas (xlsx/docx) y los
// mapfields (csv) se leen de un bucket S3 compatible (MinIO en el 207) o de
// un directorio local. Se elige con STORAGE_DRIVER=s3|local.
//
// La interfaz es la misma que tenia storageUtils con Supabase, para que los
// servicios no cambien: downloadTemplate() devuelve { success, data: Blob }
// y listTemplates() devuelve { success, templates: [{ name, metadata: { size }, updated_at }] }.
// =============================================

const fs = require('fs/promises');
const path = require('path');

const log = {
  info: (msg, data) => console.log(`[INFO] [STORAGE] ${msg}`, data !== undefined ? data : ''),
  error: (msg, data) => console.error(`[ERROR] [STORAGE] ${msg}`, data !== undefined ? data : ''),
  warn: (msg, data) => console.warn(`[WARN] [STORAGE] ${msg}`, data !== undefined ? data : '')
};

function leerConfig(env = process.env) {
  const driver = (env.STORAGE_DRIVER || 's3').toLowerCase();
  if (driver !== 's3' && driver !== 'local') {
    throw new Error(`STORAGE_DRIVER invalido: "${driver}" (usa s3 o local)`);
  }
  return {
    driver,
    bucketPlantillas: env.BUCKET_PLANTILLAS || 'plantillas-constructor',
    // Opcional: solo se usa con saveToStorage=true (hoy nadie lo pide).
    bucketGenerados: env.BUCKET_GENERADOS || '',
    plantillasDir: path.resolve(env.PLANTILLAS_DIR || './plantillas'),
    generadosDir: env.GENERADOS_DIR ? path.resolve(env.GENERADOS_DIR) : '',
    s3: {
      endpoint: env.MINIO_ENDPOINT || '',
      region: env.MINIO_REGION || 'us-east-1',
      accessKeyId: env.MINIO_ACCESS_KEY || '',
      secretAccessKey: env.MINIO_SECRET_KEY || '',
      forcePathStyle: (env.MINIO_FORCE_PATH_STYLE || 'true').toLowerCase() !== 'false',
      urlPublica: env.MINIO_URL_PUBLICA || ''
    }
  };
}

// ---------- driver S3 (MinIO, R2, AWS...) ----------
function crearDriverS3(cfg) {
  const { S3Client, GetObjectCommand, ListObjectsV2Command, PutObjectCommand } = require('@aws-sdk/client-s3');

  if (!cfg.s3.endpoint || !cfg.s3.accessKeyId || !cfg.s3.secretAccessKey) {
    throw new Error('Faltan MINIO_ENDPOINT, MINIO_ACCESS_KEY o MINIO_SECRET_KEY (STORAGE_DRIVER=s3)');
  }

  const s3 = new S3Client({
    endpoint: cfg.s3.endpoint,
    region: cfg.s3.region,
    forcePathStyle: cfg.s3.forcePathStyle,
    credentials: { accessKeyId: cfg.s3.accessKeyId, secretAccessKey: cfg.s3.secretAccessKey }
  });

  return {
    descripcion: `s3 ${cfg.s3.endpoint} bucket=${cfg.bucketPlantillas}`,

    async leer(nombre) {
      const out = await s3.send(new GetObjectCommand({ Bucket: cfg.bucketPlantillas, Key: nombre }));
      return Buffer.from(await out.Body.transformToByteArray());
    },

    async listar() {
      const objetos = [];
      let token;
      do {
        const out = await s3.send(new ListObjectsV2Command({
          Bucket: cfg.bucketPlantillas,
          ContinuationToken: token
        }));
        for (const o of out.Contents || []) {
          objetos.push({
            name: o.Key,
            metadata: { size: o.Size },
            updated_at: o.LastModified ? new Date(o.LastModified).toISOString() : null
          });
        }
        token = out.IsTruncated ? out.NextContinuationToken : undefined;
      } while (token);
      return objetos;
    },

    async subirGenerado(nombre, buffer, contentType) {
      await s3.send(new PutObjectCommand({
        Bucket: cfg.bucketGenerados,
        Key: nombre,
        Body: buffer,
        ContentType: contentType
      }));
    },

    urlGenerado(nombre) {
      const base = (cfg.s3.urlPublica || cfg.s3.endpoint).replace(/\/$/, '');
      return `${base}/${cfg.bucketGenerados}/${encodeURIComponent(nombre)}`;
    }
  };
}

// ---------- driver local (directorio montado en el contenedor) ----------
function dentroDe(dir, nombre) {
  const ruta = path.resolve(dir, nombre);
  if (ruta !== dir && !ruta.startsWith(dir + path.sep)) {
    throw new Error(`Nombre de archivo no permitido: ${nombre}`);
  }
  return ruta;
}

function crearDriverLocal(cfg) {
  return {
    descripcion: `local ${cfg.plantillasDir}`,

    async leer(nombre) {
      return fs.readFile(dentroDe(cfg.plantillasDir, nombre));
    },

    async listar() {
      const entradas = await fs.readdir(cfg.plantillasDir, { withFileTypes: true });
      const objetos = [];
      for (const e of entradas.filter((x) => x.isFile()).sort((a, b) => (a.name < b.name ? -1 : 1))) {
        const st = await fs.stat(path.join(cfg.plantillasDir, e.name));
        objetos.push({ name: e.name, metadata: { size: st.size }, updated_at: st.mtime.toISOString() });
      }
      return objetos;
    },

    async subirGenerado(nombre, buffer) {
      await fs.mkdir(cfg.generadosDir, { recursive: true });
      await fs.writeFile(dentroDe(cfg.generadosDir, nombre), buffer);
    },

    urlGenerado(nombre) {
      return `file://${dentroDe(cfg.generadosDir, nombre)}`;
    }
  };
}

function crearStorage(env = process.env) {
  const cfg = leerConfig(env);
  const driver = cfg.driver === 's3' ? crearDriverS3(cfg) : crearDriverLocal(cfg);
  const generadosConfigurado = cfg.driver === 's3' ? Boolean(cfg.bucketGenerados) : Boolean(cfg.generadosDir);

  const storageUtils = {
    descripcion: driver.descripcion,

    /**
     * Descargar plantilla o mapfield. data es un Blob (igual que supabase-js),
     * por eso excelService/wordService/mappingService no cambian.
     */
    async downloadTemplate(templateName) {
      try {
        const buffer = await driver.leer(templateName);
        return { success: true, data: new Blob([buffer]) };
      } catch (error) {
        const mensaje = error.name === 'NoSuchKey' || error.code === 'ENOENT'
          ? `Plantilla no encontrada: ${templateName}`
          : (error.message || String(error));
        log.error('Error descargando plantilla:', mensaje);
        return { success: false, error: mensaje };
      }
    },

    async listTemplates() {
      try {
        return { success: true, templates: await driver.listar() };
      } catch (error) {
        log.error('Error listando plantillas:', error.message);
        return { success: false, error: error.message };
      }
    },

    /**
     * Subir un documento generado (solo con saveToStorage=true). Sin
     * BUCKET_GENERADOS (s3) o GENERADOS_DIR (local) devuelve error y el
     * documento se entrega igual, como pasaba cuando fallaba Supabase.
     */
    async uploadGeneratedDocument(fileName, fileBuffer, metadata = {}) {
      if (!generadosConfigurado) {
        return { success: false, error: 'Almacen de documentos generados no configurado (BUCKET_GENERADOS / GENERADOS_DIR)' };
      }
      try {
        const contentType = fileName.endsWith('.docx')
          ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        await driver.subirGenerado(fileName, fileBuffer, contentType, metadata);
        return { success: true, data: { path: fileName } };
      } catch (error) {
        log.error('Error subiendo documento:', error.message);
        return { success: false, error: error.message };
      }
    },

    async getPublicUrl(fileName) {
      try {
        return { success: true, url: driver.urlGenerado(fileName) };
      } catch (error) {
        return { success: false, error: error.message };
      }
    }
  };

  return storageUtils;
}

const storageUtils = crearStorage();
log.info(`Plantillas desde ${storageUtils.descripcion}`);

module.exports = { storageUtils, crearStorage, leerConfig };
