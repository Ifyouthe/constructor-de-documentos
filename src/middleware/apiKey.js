// =============================================
// SECRETO COMPARTIDO (cabecera X-API-Key)
// Con CONSTRUCTOR_API_KEY definida, toda peticion salvo GET /health debe
// traer la cabecera X-API-Key con ese valor; si no, 401.
// Sin la variable el servicio queda ABIERTO como hasta ahora (y lo avisa al
// arrancar), para no romper a quien despliegue esta rama sin configurarla.
// =============================================

const crypto = require('crypto');

const RUTAS_LIBRES = new Set(['/health']);

function huella(valor) {
  return crypto.createHash('sha256').update(String(valor)).digest();
}

function crearApiKeyMiddleware(clave, { log = console } = {}) {
  if (!clave) {
    log.warn('CONSTRUCTOR_API_KEY no definida: el servicio acepta peticiones sin X-API-Key');
    return (req, res, next) => next();
  }

  const esperada = huella(clave);

  return (req, res, next) => {
    if (RUTAS_LIBRES.has(req.path)) return next();

    const recibida = req.get('x-api-key');
    if (recibida && crypto.timingSafeEqual(huella(recibida), esperada)) return next();

    return res.status(401).json({
      success: false,
      error: 'No autorizado: falta o no es valida la cabecera X-API-Key'
    });
  };
}

module.exports = { crearApiKeyMiddleware };
