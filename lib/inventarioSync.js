// Port directo de aplicarReceta / registrarConsumoControlUso_ / sincronizarTotalesDelDia
// de la biblioteca InventarioLib (proyecto de Apps Script aparte, que se deja intacto).
// Esto escribe en la MISMA hoja de cálculo de Inventario, así que el resultado es
// idéntico a que lo hubiera hecho la biblioteca original.
const crypto = require('crypto');
const { readAllRows, appendRow, appendRows, updateCell } = require('./sheetsHelper');
const { fechaDDMMYYYY, horaHHMM } = require('./dateHelper');

const SHEET_INVENTARIO = 'Inventario';
const SHEET_MOVIMIENTOS = 'Movimientos';
const SHEET_RECETAS = 'Recetas';
const SHEET_SYNC = 'SincronizacionDiaria';
const SHEET_CONTROL_USO = 'ControlUso';

function redondear(numero) {
  return Math.round((Number(numero) || 0) * 10000) / 10000;
}

function esVerdadero(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

async function registrarConsumoControlUso(spreadsheetId, invRows, rowIdx, sheetRow, nombreInsumo, categoria, cantidadConsumida) {
  const umbral = Number(invRows[rowIdx][13]) || 0; // N: UmbralControlUso
  if (umbral <= 0 || !cantidadConsumida) return;

  let acumulado = (Number(invRows[rowIdx][14]) || 0) + cantidadConsumida; // O: ConsumoAcumuladoControlUso
  const hoyTxt = fechaDDMMYYYY(new Date());

  let controlRows = await readAllRows(spreadsheetId, SHEET_CONTROL_USO, 'H');
  let filaAbiertaIdx = controlRows.findIndex(r =>
    String(r[1] || '').trim().toLowerCase() === nombreInsumo.toLowerCase() && !r[4]
  );

  if (filaAbiertaIdx === -1) {
    await appendRow(spreadsheetId, SHEET_CONTROL_USO, [crypto.randomUUID(), nombreInsumo, categoria, hoyTxt, '', 0, '', false]);
    controlRows = await readAllRows(spreadsheetId, SHEET_CONTROL_USO, 'H');
    filaAbiertaIdx = controlRows.length - 1;
  }

  let sheetRowControl = filaAbiertaIdx + 2;

  while (acumulado >= umbral) {
    await updateCell(spreadsheetId, SHEET_CONTROL_USO, sheetRowControl, 'F', umbral);
    await updateCell(spreadsheetId, SHEET_CONTROL_USO, sheetRowControl, 'E', hoyTxt);
    acumulado -= umbral;
    await appendRow(spreadsheetId, SHEET_CONTROL_USO, [crypto.randomUUID(), nombreInsumo, categoria, hoyTxt, '', 0, '', false]);
    controlRows = await readAllRows(spreadsheetId, SHEET_CONTROL_USO, 'H');
    sheetRowControl = controlRows.length + 1;
  }

  await updateCell(spreadsheetId, SHEET_CONTROL_USO, sheetRowControl, 'F', Math.max(acumulado, 0));
  await updateCell(spreadsheetId, SHEET_INVENTARIO, sheetRow, 'O', Math.max(acumulado, 0));
  invRows[rowIdx][14] = Math.max(acumulado, 0);
}

async function aplicarReceta(spreadsheetId, prueba, cantidadVeces) {
  cantidadVeces = Number(cantidadVeces) || 0;
  prueba = String(prueba || '').trim();
  if (!prueba) throw new Error('Falta el nombre de la prueba.');
  if (cantidadVeces === 0) {
    return { prueba, cantidadVeces: 0, aplicados: [], noEncontrados: [], sinReceta: false };
  }

  const recetaRows = await readAllRows(spreadsheetId, SHEET_RECETAS, 'C');
  const filasReceta = recetaRows.filter(r => r[0] && String(r[0]).trim().toLowerCase() === prueba.toLowerCase());

  if (!filasReceta.length) {
    return { prueba, cantidadVeces, aplicados: [], noEncontrados: [], sinReceta: true };
  }

  const invRows = await readAllRows(spreadsheetId, SHEET_INVENTARIO, 'O');
  const now = new Date();
  const fechaTxt = fechaDDMMYYYY(now);
  const horaTxt = horaHHMM(now);

  const aplicados = [];
  const noEncontrados = [];
  const movimientosNuevos = [];

  for (const fila of filasReceta) {
    const nombreInsumo = String(fila[1] || '').trim();
    const cantidadPorPrueba = Number(fila[2]) || 0;
    if (!nombreInsumo || cantidadPorPrueba <= 0) continue;

    let rowIdx = -1;
    let filaRespaldo = -1;
    for (let i = 0; i < invRows.length; i++) {
      if (String(invRows[i][1] || '').trim().toLowerCase() === nombreInsumo.toLowerCase()) {
        if (filaRespaldo === -1) filaRespaldo = i;
        if (esVerdadero(invRows[i][11])) { rowIdx = i; break; } // L: LoteActivo
      }
    }
    if (rowIdx === -1 && filaRespaldo !== -1) rowIdx = filaRespaldo;
    if (rowIdx === -1) { noEncontrados.push(nombreInsumo); continue; }

    const sheetRow = rowIdx + 2;
    const cantidadActual = Number(invRows[rowIdx][4]) || 0; // E
    const unidadesPorPieza = Number(invRows[rowIdx][8]) || 0; // I
    let unidadesRestantes = Number(invRows[rowIdx][9]) || 0; // J

    const totalConsumido = cantidadPorPrueba * cantidadVeces;
    let nuevaCantidad;

    if (unidadesPorPieza > 0) {
      if (!unidadesRestantes) unidadesRestantes = unidadesPorPieza;
      unidadesRestantes -= totalConsumido;
      let piezasConsumidas = 0;
      while (unidadesRestantes <= 0) {
        piezasConsumidas++;
        unidadesRestantes += unidadesPorPieza;
      }
      nuevaCantidad = Math.max(0, redondear(cantidadActual - piezasConsumidas));
      await updateCell(spreadsheetId, SHEET_INVENTARIO, sheetRow, 'J', unidadesRestantes);
      invRows[rowIdx][9] = unidadesRestantes;
    } else {
      nuevaCantidad = Math.max(0, redondear(cantidadActual - totalConsumido));
    }

    await updateCell(spreadsheetId, SHEET_INVENTARIO, sheetRow, 'E', nuevaCantidad);
    await updateCell(spreadsheetId, SHEET_INVENTARIO, sheetRow, 'G', now.toISOString());
    invRows[rowIdx][4] = nuevaCantidad;

    await registrarConsumoControlUso(spreadsheetId, invRows, rowIdx, sheetRow, nombreInsumo, String(invRows[rowIdx][2] || ''), totalConsumido);

    movimientosNuevos.push([crypto.randomUUID(), fechaTxt, horaTxt, nombreInsumo, totalConsumido, 'Automático', `Receta: ${cantidadVeces}x ${prueba}`]);
    aplicados.push({ insumo: nombreInsumo, cantidadDescontada: totalConsumido, nuevaCantidad });
  }

  if (movimientosNuevos.length) await appendRows(spreadsheetId, SHEET_MOVIMIENTOS, movimientosNuevos);

  return { prueba, cantidadVeces, aplicados, noEncontrados, sinReceta: false };
}

/**
 * totalesPorPrueba: { "COPROLOGICO": 16, "CPS": 34, ... } — totales de HOY ya deduplicados.
 * Aplica solo la diferencia contra lo ya sincronizado hoy, nunca duplica.
 * Nota: a diferencia del original, aquí no hay LockService entre invocaciones serverless
 * concurrentes; si dos registros casi simultáneos disparan la sincronización al mismo
 * tiempo existe una ventana de carrera teórica. Para el volumen de este laboratorio el
 * riesgo es mínimo, pero queda documentado.
 */
async function sincronizarTotalesDelDia(spreadsheetId, totalesPorPrueba) {
  if (!totalesPorPrueba || typeof totalesPorPrueba !== 'object') {
    throw new Error('Faltan los totales a sincronizar.');
  }

  const hoyTxt = fechaDDMMYYYY(new Date());
  const syncRows = await readAllRows(spreadsheetId, SHEET_SYNC, 'C');
  const resumen = [];

  for (const prueba of Object.keys(totalesPorPrueba)) {
    const nuevoTotal = Number(totalesPorPrueba[prueba]) || 0;
    let filaExistenteIdx = -1;
    let previo = 0;
    for (let i = 0; i < syncRows.length; i++) {
      if (String(syncRows[i][0] || '').trim() === hoyTxt && String(syncRows[i][1] || '').trim().toUpperCase() === prueba.toUpperCase()) {
        filaExistenteIdx = i;
        previo = Number(syncRows[i][2]) || 0;
        break;
      }
    }

    const delta = nuevoTotal - previo;
    if (delta !== 0) {
      resumen.push(await aplicarReceta(spreadsheetId, prueba, delta));
    }

    if (filaExistenteIdx === -1) {
      await appendRow(spreadsheetId, SHEET_SYNC, [hoyTxt, prueba, nuevoTotal]);
      syncRows.push([hoyTxt, prueba, nuevoTotal]);
    } else {
      await updateCell(spreadsheetId, SHEET_SYNC, filaExistenteIdx + 2, 'C', nuevoTotal);
      syncRows[filaExistenteIdx][2] = nuevoTotal;
    }
  }

  return { success: true, fecha: hoyTxt, resumen };
}

module.exports = { aplicarReceta, sincronizarTotalesDelDia };
