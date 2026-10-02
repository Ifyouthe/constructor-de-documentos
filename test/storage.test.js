const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');

process.env.STORAGE_DRIVER = 'local';
process.env.PLANTILLAS_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'plantillas-'));
const { crearStorage, leerConfig } = require('../src/config/storage');

const dir = process.env.PLANTILLAS_DIR;
fs.writeFileSync(path.join(dir, 'Mapfield de Placeholders.csv'), 'sheet,cell,placeholder,raw_text\n');
fs.writeFileSync(path.join(dir, 'seguimiento del credito (1).xlsx'), Buffer.from([1, 2, 3]));

test('local: downloadTemplate devuelve un Blob con el contenido (interfaz de supabase-js)', async () => {
  const st = crearStorage({ STORAGE_DRIVER: 'local', PLANTILLAS_DIR: dir });
  const r = await st.downloadTemplate('seguimiento del credito (1).xlsx');
  assert.equal(r.success, true);
  assert.deepEqual([...new Uint8Array(await r.data.arrayBuffer())], [1, 2, 3]);
  const csv = await st.downloadTemplate('Mapfield de Placeholders.csv');
  assert.equal(await csv.data.text(), 'sheet,cell,placeholder,raw_text\n');
});

test('local: plantilla inexistente -> success false', async () => {
  const st = crearStorage({ STORAGE_DRIVER: 'local', PLANTILLAS_DIR: dir });
  const r = await st.downloadTemplate('no-existe.xlsx');
  assert.equal(r.success, false);
  assert.match(r.error, /no encontrada/);
});

test('local: no deja salir del directorio', async () => {
  const st = crearStorage({ STORAGE_DRIVER: 'local', PLANTILLAS_DIR: dir });
  for (const nombre of ['../../etc/passwd', '/etc/passwd']) {
    const r = await st.downloadTemplate(nombre);
    assert.equal(r.success, false, nombre);
  }
});

test('local: listTemplates con la forma de supabase (name, metadata.size, updated_at)', async () => {
  const st = crearStorage({ STORAGE_DRIVER: 'local', PLANTILLAS_DIR: dir });
  const r = await st.listTemplates();
  assert.equal(r.success, true);
  assert.deepEqual(r.templates.map((t) => t.name), ['Mapfield de Placeholders.csv', 'seguimiento del credito (1).xlsx']);
  assert.equal(r.templates[1].metadata.size, 3);
  assert.ok(r.templates[1].updated_at);
});

test('generados sin configurar -> error controlado (el documento se entrega igual)', async () => {
  const st = crearStorage({ STORAGE_DRIVER: 'local', PLANTILLAS_DIR: dir });
  const r = await st.uploadGeneratedDocument('x.xlsx', Buffer.from('x'));
  assert.equal(r.success, false);
});

test('s3 sin credenciales -> falla al arrancar con mensaje claro', () => {
  assert.throws(() => crearStorage({ STORAGE_DRIVER: 's3' }), /MINIO_ENDPOINT/);
  assert.throws(() => leerConfig({ STORAGE_DRIVER: 'supabase' }), /STORAGE_DRIVER invalido/);
});

// Contra un MinIO real (opcional): TEST_MINIO_ENDPOINT, TEST_MINIO_ACCESS_KEY, TEST_MINIO_SECRET_KEY, TEST_BUCKET_PLANTILLAS
test('s3 contra MinIO real', { skip: !process.env.TEST_MINIO_ENDPOINT && 'sin TEST_MINIO_ENDPOINT' }, async () => {
  const st = crearStorage({
    STORAGE_DRIVER: 's3',
    MINIO_ENDPOINT: process.env.TEST_MINIO_ENDPOINT,
    MINIO_ACCESS_KEY: process.env.TEST_MINIO_ACCESS_KEY,
    MINIO_SECRET_KEY: process.env.TEST_MINIO_SECRET_KEY,
    BUCKET_PLANTILLAS: process.env.TEST_BUCKET_PLANTILLAS || 'plantillas-constructor'
  });
  const lista = await st.listTemplates();
  assert.equal(lista.success, true);
  assert.ok(lista.templates.length > 0);
  const primero = lista.templates[0];
  const r = await st.downloadTemplate(primero.name);
  assert.equal(r.success, true);
  assert.equal((await r.data.arrayBuffer()).byteLength, primero.metadata.size);
  assert.equal((await st.downloadTemplate('no-existe.xlsx')).success, false);
});
