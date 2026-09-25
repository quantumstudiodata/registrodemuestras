// Monterrey y Ciudad de México ya no usan horario de verano (desde 2022), así que un
// desfase fijo de -6h reproduce exactamente lo que hacía getNow() en el Code.gs original.
const OFFSET_MS = -6 * 60 * 60 * 1000;

function partsMX(date) {
  const shifted = new Date((date || new Date()).getTime() + OFFSET_MS);
  return {
    yyyy: shifted.getUTCFullYear(),
    MM: String(shifted.getUTCMonth() + 1).padStart(2, '0'),
    dd: String(shifted.getUTCDate()).padStart(2, '0'),
    HH: String(shifted.getUTCHours()).padStart(2, '0'),
    mm: String(shifted.getUTCMinutes()).padStart(2, '0'),
    ss: String(shifted.getUTCSeconds()).padStart(2, '0')
  };
}

function getNow(date) {
  const p = partsMX(date);
  return { fecha: `${p.yyyy}-${p.MM}-${p.dd}`, hora: `${p.HH}:${p.mm}:${p.ss}` };
}

function fechaDDMMYYYY(date) {
  const p = partsMX(date);
  return `${p.dd}/${p.MM}/${p.yyyy}`;
}

function horaHHMM(date) {
  const p = partsMX(date);
  return `${p.HH}:${p.mm}`;
}

// Filas guardadas antes de la migración (por el Code.gs original) quedaron con la
// columna Fecha/Hora como un valor de fecha real de Sheets (Google la auto-convirtió al
// detectar el patrón), no como texto plano. Con valueRenderOption UNFORMATTED_VALUE esas
// celdas llegan como número serial de Sheets; las filas nuevas (escritas con RAW) llegan
// como texto 'yyyy-MM-dd'/'HH:mm:ss' de siempre. Estas funciones normalizan ambos casos.
const SHEETS_EPOCH_UTC_MS = Date.UTC(1899, 11, 30);

function normalizeFechaCelda(valor) {
  if (valor === '' || valor === null || valor === undefined) return '';
  if (typeof valor === 'number') {
    const ms = SHEETS_EPOCH_UTC_MS + Math.floor(valor) * 86400000;
    const d = new Date(ms);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  return String(valor).substring(0, 10);
}

function normalizeHoraCelda(valor) {
  if (valor === '' || valor === null || valor === undefined) return '';
  if (typeof valor === 'number') {
    const totalSeg = Math.round((valor % 1) * 86400);
    const hh = Math.floor(totalSeg / 3600);
    const mm = Math.floor((totalSeg % 3600) / 60);
    const ss = totalSeg % 60;
    return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}:${String(ss).padStart(2, '0')}`;
  }
  return String(valor).substring(0, 8);
}

module.exports = { getNow, fechaDDMMYYYY, horaHHMM, normalizeFechaCelda, normalizeHoraCelda };
