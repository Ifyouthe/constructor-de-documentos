// Prueba de punta a punta del servidor con el driver local y una plantilla
// sintetica (las plantillas reales de Sumate NO van en este repo publico).
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const net = require('net');
const { spawn } = require('child_process');
const ExcelJS = require('exceljs');

const CLAVE = 'clave-de-prueba';

function puertoLibre() {
  return new Promise((ok) => {
    const s = net.createServer().listen(0, () => { const p = s.address().port; s.close(() => ok(p)); });
  });
}

async function plantillaSintetica(dir) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Scoring del Cliente');
  ws.getCell('A4').value = 'BC Score';
  ws.getCell('B4').value = '{buro.BC_score}';
  await wb.xlsx.writeFile(path.join(dir, 'SCORING_CON_HC.xlsx'));
  fs.writeFileSync(path.join(dir, 'Mapfield_Con_HC.csv'),
    'sheet,cell,placeholder,raw_text\nScoring del Cliente,B4,buro.BC_score,{buro.BC_score}\n');
}

test('servidor: secreto, limite de tamano y generacion con driver local', async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'plantillas-'));
  await plantillaSintetica(dir);
  const port = await puertoLibre();
  const hijo = spawn(process.execPath, ['server.js'], {
    cwd: path.join(__dirname, '..'),
    env: { ...process.env, PORT: String(port), STORAGE_DRIVER: 'local', PLANTILLAS_DIR: dir, CONSTRUCTOR_API_KEY: CLAVE, MAX_BODY_SIZE: '10kb', LOG_LEVEL: 'ERROR' },
    stdio: 'ignore'
  });
  t.after(() => hijo.kill());
  const base = `http://127.0.0.1:${port}`;

  // Espera a que arranque (hasta 30 s: en una maquina cargada tarda)
  let salud;
  for (let i = 0; i < 300 && !salud; i++) {
    try { salud = await fetch(`${base}/health`); } catch { await new Promise((r) => setTimeout(r, 100)); }
  }
  assert.ok(salud, 'el servidor no arranco en 30 s');
  assert.equal(salud.status, 200, '/health sin clave responde 200');
  const cuerpoSalud = await salud.json();
  assert.equal(cuerpoSalud.checks.storage, 'healthy');
  assert.equal(cuerpoSalud.checks.templatesCount, 2);

  const payload = JSON.stringify({ formato: 'con_HC', 'buro.BC_score': '712', codigo_de_prospecto: 'PRUEBA-1' });
  const post = (headers, body = payload) => fetch(`${base}/webhook/generar-documento`, {
    method: 'POST', headers: { 'Content-Type': 'application/json', ...headers }, body
  });

  assert.equal((await post({})).status, 401, 'sin X-API-Key');
  assert.equal((await post({ 'X-API-Key': 'mala' })).status, 401, 'X-API-Key incorrecta');
  assert.equal((await fetch(`${base}/api/plantillas`)).status, 401, 'tambien /api/* pide clave');

  const ok = await post({ 'X-API-Key': CLAVE });
  assert.equal(ok.status, 200);
  assert.match(ok.headers.get('content-type'), /spreadsheetml/);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(await ok.arrayBuffer()));
  assert.equal(wb.getWorksheet('Scoring del Cliente').getCell('B4').value, '712');

  const grande = JSON.stringify({ formato: 'con_HC', relleno: 'x'.repeat(20 * 1024) });
  assert.equal((await post({ 'X-API-Key': CLAVE }, grande)).status, 413, 'cuerpo mayor que MAX_BODY_SIZE');

  for (const ruta of ['/api/stats', '/api/documentos', '/api/descargar/1']) {
    assert.equal((await fetch(`${base}${ruta}`, { headers: { 'X-API-Key': CLAVE } })).status, 404, `${ruta} retirada (dependia de Supabase)`);
  }
});
