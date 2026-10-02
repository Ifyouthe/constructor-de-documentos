#!/usr/bin/env node
// =============================================
// Copia las plantillas del bucket de Supabase (documentossumate, en
// supabase-evolve) al bucket S3/MinIO del constructor y/o a un directorio.
//
// En Supabase SOLO LEE (list + download). Verifica cada archivo por sha256
// releyendolo del destino. No imprime credenciales.
//
// Uso:
//   SUPABASE_URL=... SUPABASE_KEY=... [SUPABASE_BUCKET=documentossumate] \
//   MINIO_ENDPOINT=... MINIO_ACCESS_KEY=... MINIO_SECRET_KEY=... [BUCKET_PLANTILLAS=plantillas-constructor] \
//   node scripts/copiar-plantillas.js [--s3] [--dir ./plantillas] [--dry-run]
//
// Tambien sirve para subir a S3 una carpeta ya descargada:
//   node scripts/copiar-plantillas.js --desde-dir ./plantillas --s3
// =============================================

const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');

function args(argv) {
  const out = { s3: false, dir: '', desdeDir: '', dryRun: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--s3') out.s3 = true;
    else if (a === '--dry-run') out.dryRun = true;
    else if (a === '--dir') out.dir = argv[++i];
    else if (a === '--desde-dir') out.desdeDir = argv[++i];
    else throw new Error(`Argumento desconocido: ${a}`);
  }
  if (!out.s3 && !out.dir && !out.dryRun) throw new Error('Indica al menos un destino: --s3 y/o --dir <ruta> (o --dry-run)');
  return out;
}

const sha256 = (buf) => crypto.createHash('sha256').update(buf).digest('hex');

// ---------- origen: Supabase Storage (REST, solo lectura) ----------
function origenSupabase(env) {
  const url = (env.SUPABASE_URL || '').replace(/\/$/, '');
  const key = env.SUPABASE_KEY || env.SUPABASE_ANON_KEY || env.SUPABASE_SERVICE_ROLE_KEY;
  const bucket = env.SUPABASE_BUCKET || 'documentossumate';
  if (!url || !key) throw new Error('Faltan SUPABASE_URL y SUPABASE_KEY');
  const headers = { apikey: key, Authorization: `Bearer ${key}` };

  async function listar(prefix = '') {
    const res = await fetch(`${url}/storage/v1/object/list/${bucket}`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json' },
      body: JSON.stringify({ prefix, limit: 1000, offset: 0, sortBy: { column: 'name', order: 'asc' } })
    });
    if (!res.ok) throw new Error(`list ${bucket}/${prefix}: HTTP ${res.status}`);
    const items = await res.json();
    const nombres = [];
    for (const it of items) {
      const nombre = prefix ? `${prefix}/${it.name}` : it.name;
      if (it.id === null) nombres.push(...(await listar(nombre))); // carpeta
      else if (it.name !== '.emptyFolderPlaceholder') nombres.push(nombre);
    }
    return nombres;
  }

  return {
    descripcion: `supabase bucket=${bucket}`,
    listar,
    async leer(nombre) {
      const ruta = nombre.split('/').map(encodeURIComponent).join('/');
      const res = await fetch(`${url}/storage/v1/object/${bucket}/${ruta}`, { headers });
      if (!res.ok) throw new Error(`download ${nombre}: HTTP ${res.status}`);
      return Buffer.from(await res.arrayBuffer());
    }
  };
}

function origenDir(dir) {
  return {
    descripcion: `dir ${dir}`,
    async listar() {
      return (await fs.readdir(dir, { withFileTypes: true })).filter((e) => e.isFile()).map((e) => e.name).sort();
    },
    leer: (nombre) => fs.readFile(path.join(dir, nombre))
  };
}

// ---------- destino: S3/MinIO ----------
function destinoS3(env) {
  const {
    S3Client, CreateBucketCommand, HeadBucketCommand, PutObjectCommand, GetObjectCommand
  } = require('@aws-sdk/client-s3');
  if (!env.MINIO_ENDPOINT || !env.MINIO_ACCESS_KEY || !env.MINIO_SECRET_KEY) {
    throw new Error('Faltan MINIO_ENDPOINT, MINIO_ACCESS_KEY o MINIO_SECRET_KEY');
  }
  const bucket = env.BUCKET_PLANTILLAS || 'plantillas-constructor';
  const s3 = new S3Client({
    endpoint: env.MINIO_ENDPOINT,
    region: env.MINIO_REGION || 'us-east-1',
    forcePathStyle: (env.MINIO_FORCE_PATH_STYLE || 'true').toLowerCase() !== 'false',
    credentials: { accessKeyId: env.MINIO_ACCESS_KEY, secretAccessKey: env.MINIO_SECRET_KEY }
  });
  return {
    descripcion: `s3 bucket=${bucket}`,
    async preparar() {
      try {
        await s3.send(new HeadBucketCommand({ Bucket: bucket }));
      } catch {
        await s3.send(new CreateBucketCommand({ Bucket: bucket }));
        console.log(`Bucket creado: ${bucket}`);
      }
    },
    async escribir(nombre, buf) {
      await s3.send(new PutObjectCommand({ Bucket: bucket, Key: nombre, Body: buf }));
      const out = await s3.send(new GetObjectCommand({ Bucket: bucket, Key: nombre }));
      return Buffer.from(await out.Body.transformToByteArray());
    }
  };
}

function destinoDir(dir) {
  return {
    descripcion: `dir ${dir}`,
    preparar: () => fs.mkdir(dir, { recursive: true }),
    async escribir(nombre, buf) {
      const ruta = path.join(dir, nombre);
      await fs.mkdir(path.dirname(ruta), { recursive: true });
      await fs.writeFile(ruta, buf);
      return fs.readFile(ruta);
    }
  };
}

async function main() {
  require('dotenv').config();
  const opts = args(process.argv.slice(2));
  const origen = opts.desdeDir ? origenDir(path.resolve(opts.desdeDir)) : origenSupabase(process.env);
  const destinos = [];
  if (opts.s3) destinos.push(destinoS3(process.env));
  if (opts.dir) destinos.push(destinoDir(path.resolve(opts.dir)));

  console.log(`Origen: ${origen.descripcion} | Destinos: ${destinos.map((d) => d.descripcion).join(', ') || '(ninguno, dry-run)'}`);
  if (!opts.dryRun) for (const d of destinos) await d.preparar();

  const nombres = await origen.listar();
  let bytes = 0;
  let fallos = 0;
  for (const nombre of nombres) {
    const buf = await origen.leer(nombre);
    const hash = sha256(buf);
    bytes += buf.length;
    let estado = 'leido';
    if (!opts.dryRun) {
      for (const d of destinos) {
        const copia = await d.escribir(nombre, buf);
        if (sha256(copia) !== hash) {
          estado = `DISTINTO en ${d.descripcion}`;
          fallos++;
        } else if (estado === 'leido') {
          estado = 'copiado y verificado';
        }
      }
    }
    console.log(`${hash.slice(0, 12)}  ${String(buf.length).padStart(8)}  ${nombre}  [${estado}]`);
  }
  console.log(`Total: ${nombres.length} archivos, ${bytes} bytes, ${fallos} fallos`);
  if (fallos) process.exit(1);
}

main().catch((e) => {
  console.error(`ERROR: ${e.message}`);
  process.exit(1);
});
