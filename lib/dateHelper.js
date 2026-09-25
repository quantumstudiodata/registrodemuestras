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

module.exports = { getNow, fechaDDMMYYYY, horaHHMM };
