const test = require('node:test');
const assert = require('node:assert');
const { crearApiKeyMiddleware } = require('../src/middleware/apiKey');

const silencioso = { warn() {} };

function llamar(mw, { path = '/webhook/generar-documento', clave } = {}) {
  const req = { path, get: (h) => (h.toLowerCase() === 'x-api-key' ? clave : undefined) };
  const res = {
    statusCode: 200,
    body: null,
    status(c) { this.statusCode = c; return this; },
    json(b) { this.body = b; return this; }
  };
  let siguio = false;
  mw(req, res, () => { siguio = true; });
  return { siguio, res };
}

test('sin CONSTRUCTOR_API_KEY el servicio queda abierto (compatibilidad)', () => {
  const mw = crearApiKeyMiddleware('', { log: silencioso });
  assert.equal(llamar(mw).siguio, true);
});

test('con clave: sin cabecera -> 401', () => {
  const mw = crearApiKeyMiddleware('clave-de-prueba', { log: silencioso });
  const r = llamar(mw);
  assert.equal(r.siguio, false);
  assert.equal(r.res.statusCode, 401);
});

test('con clave: cabecera incorrecta (incluso de otro largo) -> 401', () => {
  const mw = crearApiKeyMiddleware('clave-de-prueba', { log: silencioso });
  assert.equal(llamar(mw, { clave: 'otra' }).res.statusCode, 401);
  assert.equal(llamar(mw, { clave: 'clave-de-prueba-x' }).res.statusCode, 401);
});

test('con clave: cabecera correcta -> pasa', () => {
  const mw = crearApiKeyMiddleware('clave-de-prueba', { log: silencioso });
  assert.equal(llamar(mw, { clave: 'clave-de-prueba' }).siguio, true);
});

test('/health no pide clave (healthcheck de Docker)', () => {
  const mw = crearApiKeyMiddleware('clave-de-prueba', { log: silencioso });
  assert.equal(llamar(mw, { path: '/health' }).siguio, true);
});
