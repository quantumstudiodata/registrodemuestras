// Port directo de Code.gs (Registro de Muestras) a Node, usando la API de Sheets con
// una cuenta de servicio en vez de SpreadsheetApp. Misma lógica de negocio, mismas
// columnas (A-J: ID, Fecha, Hora, Codigo, Prueba, Analista, Nombre, Ronda, Sucursal, Estado).
const crypto = require('crypto');
const {
  ensureSheet,
  ensureMonthSheet,
  getAllMonthSheetNames,
  getSheetIdByTitle,
  readAllRows,
  appendRow,
  appendRows,
  updateCell,
  deleteRow,
  deleteRows
} = require('./sheetsHelper');
const { getNow } = require('./dateHelper');
const { subirFotoPerfil } = require('./drive');
const { sincronizarTotalesDelDia } = require('./inventarioSync');

const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID_MUESTRAS;
const SHEET_PERFILES = 'Perfiles';
const HEADER_PERFILES = ['Usuario', 'NombreCompleto', 'FotoURL'];
const SHEET_MURO = 'Muro';
const HEADER_MURO = ['ID', 'Fecha', 'Hora', 'Usuario', 'Mensaje'];

function requireSpreadsheetId() {
  if (!SPREADSHEET_ID) throw new Error('Falta la variable de entorno GOOGLE_SHEET_ID_MUESTRAS.');
}

function rowToMuestra(row) {
  const fecha = String(row[1] || '').substring(0, 10);
  const hora = String(row[2] || '').substring(0, 8);
  const estado = String(row[9] || '');
  const prueba = String(row[4] || '');
  return {
    id: String(row[0]), fecha, hora, codigo: String(row[3] || ''),
    prueba: prueba === 'RECHAZADA' ? '' : prueba,
    analista: String(row[5] || ''), nombre: String(row[6] || ''),
    ronda: Number(row[7] || 1), sucursal: String(row[8] || ''),
    rechazada: prueba === 'RECHAZADA', manual: (estado === 'manual' || estado === 'auto')
  };
}

async function saveMuestra(codigo, pruebas, analista, nombre, ronda) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, t.fecha);
    const ids = [];
    const lista = (Array.isArray(pruebas) && pruebas.length > 0) ? pruebas : [''];
    const rows = lista.map(prueba => {
      const id = crypto.randomUUID();
      ids.push(id);
      return [id, t.fecha, t.hora, String(codigo), String(prueba || ''), String(analista || ''), String(nombre || ''), Number(ronda || 1), '', ''];
    });
    await appendRows(SPREADSHEET_ID, sheetTitle, rows);
    return { success: true, ids, fecha: t.fecha, hora: t.hora };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function getMuestrasHoy() {
  try {
    requireSpreadsheetId();
    const hojas = await getAllMonthSheetNames(SPREADSHEET_ID);
    const rows = [];
    for (const sheetTitle of hojas) {
      const data = await readAllRows(SPREADSHEET_ID, sheetTitle);
      data.filter(r => r[0]).forEach(r => rows.push(rowToMuestra(r)));
    }
    return { success: true, muestras: rows };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function actualizarPruebas(actualizaciones) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const lista = Array.isArray(actualizaciones) ? actualizaciones : [actualizaciones];

    for (const upd of lista) {
      if (upd.auto === true || upd.auto === 'true') {
        const fechaAuto = String(upd.fecha || t.fecha);
        const shAuto = await ensureMonthSheet(SPREADSHEET_ID, fechaAuto);
        await appendRow(SPREADSHEET_ID, shAuto, [crypto.randomUUID(), fechaAuto, t.hora, String(upd.codigo || ''), String(upd.prueba || ''), String(upd.analista || ''), '', 1, String(upd.sucursal || ''), 'auto']);
        continue;
      }

      if (upd.manual === 'true' || upd.manual === true) {
        const shManual = await ensureMonthSheet(SPREADSHEET_ID, t.fecha);
        await appendRow(SPREADSHEET_ID, shManual, [crypto.randomUUID(), t.fecha, t.hora, String(upd.codigo || ''), String(upd.prueba || ''), String(upd.analista || ''), '', 1, String(upd.sucursal || ''), 'manual']);
        continue;
      }

      const fechaRegistro = String(upd.fecha || t.fecha);
      const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, fechaRegistro);
      const data = await readAllRows(SPREADSHEET_ID, sheetTitle);

      if (upd.rechazada === true || upd.rechazada === 'true') {
        const idx = data.findIndex(r => String(r[0]) === String(upd.id));
        if (idx !== -1) {
          const sheetRow = idx + 2;
          await updateCell(SPREADSHEET_ID, sheetTitle, sheetRow, 'E', 'RECHAZADA');
          if (upd.razon) await updateCell(SPREADSHEET_ID, sheetTitle, sheetRow, 'G', String(upd.razon));
        }
        continue;
      }

      const idx2 = data.findIndex(r => String(r[0]) === String(upd.id));
      if (idx2 !== -1) {
        const sheetRow = idx2 + 2;
        const pruebaActual = String(data[idx2][4] || '');
        if (pruebaActual === '' || pruebaActual === 'undefined') {
          if (upd.prueba !== undefined) await updateCell(SPREADSHEET_ID, sheetTitle, sheetRow, 'E', upd.prueba);
          if (upd.sucursal) await updateCell(SPREADSHEET_ID, sheetTitle, sheetRow, 'I', upd.sucursal);
        }
      }
    }

    await sincronizarInventarioHoy();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function deleteMuestra(id, fecha) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, String(fecha || t.fecha));
    const data = await readAllRows(SPREADSHEET_ID, sheetTitle);
    const idx = data.findIndex(r => String(r[0]) === String(id));
    if (idx === -1) return { success: false, error: 'No encontrada' };
    await deleteRow(SPREADSHEET_ID, sheetTitle, idx + 2);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function guardarManual(payload) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, String(payload.fecha || t.fecha));
    const id = crypto.randomUUID();
    await appendRow(SPREADSHEET_ID, sheetTitle, [id, String(payload.fecha || ''), String(payload.hora || ''), String(payload.codigo || ''), String(payload.prueba || ''), String(payload.analista || ''), '', 1, String(payload.sucursal || ''), 'manual']);
    await sincronizarInventarioHoy();
    return { success: true, id };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function deleteByCode(codigo) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, t.fecha);
    const data = await readAllRows(SPREADSHEET_ID, sheetTitle);
    const filasABorrar = [];
    for (let r = 0; r < data.length; r++) {
      const fecha = String(data[r][1] || '').substring(0, 10);
      if (String(data[r][3]) === String(codigo) && fecha === t.fecha) filasABorrar.push(r + 2);
    }
    await deleteRows(SPREADSHEET_ID, sheetTitle, filasABorrar);
    return { success: true, deleted: filasABorrar.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function getMuestrasMes(mes) {
  try {
    requireSpreadsheetId();
    const sheetId = await getSheetIdByTitle(SPREADSHEET_ID, mes);
    if (sheetId === null) return { success: true, muestras: [] };
    const data = await readAllRows(SPREADSHEET_ID, mes);
    return { success: true, muestras: data.filter(r => r[0]).map(rowToMuestra) };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function deleteMuestraGrupo(codigo, fecha) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, String(fecha || t.fecha));
    const data = await readAllRows(SPREADSHEET_ID, sheetTitle);
    const filasABorrar = [];
    for (let r = 0; r < data.length; r++) {
      const fechaFila = String(data[r][1] || '').substring(0, 10);
      if (String(data[r][3]) === String(codigo) && fechaFila === String(fecha)) filasABorrar.push(r + 2);
    }
    await deleteRows(SPREADSHEET_ID, sheetTitle, filasABorrar);
    return { success: true, eliminadas: filasABorrar.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function deleteFilaPorId(id, fecha) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, String(fecha || t.fecha));
    const data = await readAllRows(SPREADSHEET_ID, sheetTitle);
    const idx = data.findIndex(r => String(r[0]) === String(id));
    if (idx === -1) return { success: false, error: 'No encontrada' };
    await deleteRow(SPREADSHEET_ID, sheetTitle, idx + 2);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function editarPruebaFila(id, fecha, nuevaPrueba) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, String(fecha || t.fecha));
    const data = await readAllRows(SPREADSHEET_ID, sheetTitle);
    const idx = data.findIndex(r => String(r[0]) === String(id));
    if (idx === -1) return { success: false, error: 'No encontrada' };
    await updateCell(SPREADSHEET_ID, sheetTitle, idx + 2, 'E', String(nuevaPrueba || ''));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function renombrarCodigo(codigoViejo, codigoNuevo, fecha, sucursalNueva) {
  try {
    requireSpreadsheetId();
    const t = getNow();
    const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, String(fecha || t.fecha));
    const data = await readAllRows(SPREADSHEET_ID, sheetTitle);
    let actualizadas = 0;
    for (let r = 0; r < data.length; r++) {
      const fechaFila = String(data[r][1] || '').substring(0, 10);
      if (String(data[r][3]) === String(codigoViejo) && fechaFila === String(fecha)) {
        await updateCell(SPREADSHEET_ID, sheetTitle, r + 2, 'D', String(codigoNuevo));
        if (sucursalNueva) await updateCell(SPREADSHEET_ID, sheetTitle, r + 2, 'I', String(sucursalNueva));
        actualizadas++;
      }
    }
    return { success: true, actualizadas };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function getMuestrasMeses(meses) {
  try {
    requireSpreadsheetId();
    const rows = [];
    for (const mes of meses) {
      const sheetId = await getSheetIdByTitle(SPREADSHEET_ID, mes);
      if (sheetId === null) continue;
      const data = await readAllRows(SPREADSHEET_ID, mes);
      data.filter(r => r[0]).forEach(r => rows.push(rowToMuestra(r)));
    }
    return { success: true, muestras: rows };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ---------- Perfiles ----------

async function getPerfil(usuario) {
  try {
    requireSpreadsheetId();
    await ensureSheet(SPREADSHEET_ID, SHEET_PERFILES, HEADER_PERFILES);
    const data = await readAllRows(SPREADSHEET_ID, SHEET_PERFILES, 'C');
    const found = data.find(r => String(r[0]) === String(usuario));
    if (found) return { success: true, nombreCompleto: String(found[1] || ''), fotoURL: String(found[2] || '') };
    return { success: true, nombreCompleto: '', fotoURL: '' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function guardarPerfil(usuario, nombreCompleto, fotoBase64) {
  try {
    requireSpreadsheetId();
    await ensureSheet(SPREADSHEET_ID, SHEET_PERFILES, HEADER_PERFILES);
    const data = await readAllRows(SPREADSHEET_ID, SHEET_PERFILES, 'C');
    let fotoURL = '';
    if (fotoBase64) fotoURL = await subirFotoPerfil(usuario, fotoBase64);

    const idx = data.findIndex(r => String(r[0]) === String(usuario));
    if (idx !== -1) {
      await updateCell(SPREADSHEET_ID, SHEET_PERFILES, idx + 2, 'B', nombreCompleto || '');
      if (fotoURL) await updateCell(SPREADSHEET_ID, SHEET_PERFILES, idx + 2, 'C', fotoURL);
      return { success: true, fotoURL: fotoURL || String(data[idx][2] || '') };
    }
    await appendRow(SPREADSHEET_ID, SHEET_PERFILES, [usuario, nombreCompleto || '', fotoURL]);
    return { success: true, fotoURL };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ---------- Muro ----------

async function getMuro() {
  try {
    requireSpreadsheetId();
    await ensureSheet(SPREADSHEET_ID, SHEET_MURO, HEADER_MURO);
    const data = await readAllRows(SPREADSHEET_ID, SHEET_MURO, 'E');
    const mensajes = data.filter(r => r[0]).map(r => ({
      id: String(r[0]), fecha: String(r[1] || ''), hora: String(r[2] || ''),
      usuario: String(r[3] || ''), mensaje: String(r[4] || '')
    }));
    return { success: true, mensajes };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function publicarMensaje(usuario, mensaje) {
  try {
    requireSpreadsheetId();
    await ensureSheet(SPREADSHEET_ID, SHEET_MURO, HEADER_MURO);
    const t = getNow();
    const id = crypto.randomUUID();
    await appendRow(SPREADSHEET_ID, SHEET_MURO, [id, t.fecha, t.hora, usuario, mensaje]);
    return { success: true, id, fecha: t.fecha, hora: t.hora };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function eliminarMensajeMuro(id) {
  try {
    requireSpreadsheetId();
    const data = await readAllRows(SPREADSHEET_ID, SHEET_MURO, 'E');
    const idx = data.findIndex(r => String(r[0]) === String(id));
    if (idx === -1) return { success: false, error: 'No encontrado' };
    await deleteRow(SPREADSHEET_ID, SHEET_MURO, idx + 2);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ---------- Estadísticas + sincronización de inventario ----------

const PRUEBAS_ESTADISTICAS = {
  'COPROLOGICO': ['COPROLOGICO', 'CPS', 'SOH', 'AMF'],
  'HELIH': ['HELIH'],
  'CPS': ['CPS'], 'CPS2': ['CPS'], 'CPS3': ['CPS'],
  'SOH': ['SOH'], 'SOH2': ['SOH'], 'SOH3': ['SOH'],
  'AMF': ['AMF'], 'AMF2': ['AMF'], 'AMF3': ['AMF'],
  'ROTA_ADENO_ASTRO': ['ROTA_ADENO_ASTRO'],
  'CALPROT_LACTO': ['CALPROT_LACTO'],
  'FOB_TRANSF': ['FOB_TRANSF'],
  'ENTAMOEBA': ['ENTAMOEBA'],
  'CRYPTO_GIARDIA': ['CRYPTO_GIARDIA'],
  'CLOSTRIDIUM': ['CLOSTRIDIUM'],
  'GRAM': ['GRAM'],
  'WRIGHT': ['WRIGHT']
};

function extraerCodigoPaciente(cod) {
  const s = String(cod || '').split(/[-']/)[0];
  const r = s.length > 2 ? s.substring(2) : s;
  return r.replace(/\.0$/, '');
}

async function getStatsHoyServidor() {
  const t = getNow();
  const sheetTitle = await ensureMonthSheet(SPREADSHEET_ID, t.fecha);
  const data = await readAllRows(SPREADSHEET_ID, sheetTitle);
  const counts = {};
  const vistos = {};

  for (const row of data) {
    if (!row[0]) continue;
    const fecha = String(row[1] || '').substring(0, 10);
    if (fecha !== t.fecha) continue;
    const prueba = String(row[4] || '');
    if (!prueba || prueba === 'RECHAZADA') continue;
    const cats = PRUEBAS_ESTADISTICAS[prueba];
    if (!cats) continue;
    const codigoPac = extraerCodigoPaciente(row[3]);
    cats.forEach(cat => {
      const clave = `${codigoPac}_${prueba}_${cat}`;
      if (!vistos[clave]) { vistos[clave] = true; counts[cat] = (counts[cat] || 0) + 1; }
    });
  }
  return counts;
}

// A diferencia del original (llamada directa a la Biblioteca InventarioLib), aquí se
// escribe directo contra la hoja de Inventario con el mismo service account, sin tocar
// el proyecto de Apps Script de Inventario. Si GOOGLE_SHEET_ID_INVENTARIO no está
// configurada, la sincronización simplemente se omite (no rompe el registro de muestras).
async function sincronizarInventarioHoy() {
  try {
    const inventarioId = process.env.GOOGLE_SHEET_ID_INVENTARIO;
    if (!inventarioId) return null;
    const stats = await getStatsHoyServidor();
    return await sincronizarTotalesDelDia(inventarioId, stats);
  } catch (err) {
    console.error('No se pudo sincronizar el inventario:', err.message);
    return null;
  }
}

module.exports = {
  getNow,
  saveMuestra,
  getMuestrasHoy,
  actualizarPruebas,
  deleteMuestra,
  guardarManual,
  deleteByCode,
  getMuestrasMes,
  deleteMuestraGrupo,
  deleteFilaPorId,
  editarPruebaFila,
  renombrarCodigo,
  getMuestrasMeses,
  getPerfil,
  guardarPerfil,
  getMuro,
  publicarMensaje,
  eliminarMensajeMuro
};
