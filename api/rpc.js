// Único endpoint que reemplaza a google.script.run: recibe { fn, args } y despacha a la
// función correspondiente en lib/muestras.js. El shim del lado del cliente (en index.html)
// llama a esto y reparte la respuesta a withSuccessHandler/withFailureHandler exactamente
// como lo hacía Apps Script.
const muestras = require('../lib/muestras');

async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ message: 'Método no permitido' });
    return;
  }

  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch (e) { body = {}; }
  }
  const { fn, args } = body || {};

  const target = muestras[fn];
  if (typeof target !== 'function') {
    res.status(400).json({ message: `Función desconocida: ${fn}` });
    return;
  }

  try {
    const result = await target(...(Array.isArray(args) ? args : []));
    res.status(200).json(result);
  } catch (err) {
    console.error(`Error en ${fn}:`, err);
    res.status(500).json({ message: err.message || 'Error interno del servidor.' });
  }
}

// Sube el límite del body para las fotos de perfil en base64 (el límite real de la
// plataforma en el plan Hobby de Vercel sigue siendo ~4.5MB por request).
handler.config = {
  api: { bodyParser: { sizeLimit: '10mb' } }
};

module.exports = handler;
