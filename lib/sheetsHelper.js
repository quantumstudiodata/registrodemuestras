const { getSheetsClient } = require('./auth');

const HEADER_MUESTRAS = ['ID', 'Fecha', 'Hora', 'Codigo', 'Prueba', 'Analista', 'Nombre', 'Ronda', 'Sucursal', 'Estado'];

async function getSheetProps(spreadsheetId) {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  return res.data.sheets.map(s => s.properties);
}

async function getSheetIdByTitle(spreadsheetId, title) {
  const props = await getSheetProps(spreadsheetId);
  const found = props.find(p => p.title === title);
  return found ? found.sheetId : null;
}

function getMonthSheetName(fecha) {
  return String(fecha).substring(0, 7);
}

async function ensureSheet(spreadsheetId, title, header) {
  const props = await getSheetProps(spreadsheetId);
  const found = props.find(p => p.title === title);
  if (found) return found.sheetId;

  const sheets = await getSheetsClient();
  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: { title } } }] }
  });
  const sheetId = addRes.data.replies[0].addSheet.properties.sheetId;

  if (header && header.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${title}'!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: [header] }
    });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: { userEnteredFormat: { textFormat: { bold: true } } },
            fields: 'userEnteredFormat.textFormat.bold'
          }
        }]
      }
    });
  }
  return sheetId;
}

async function ensureMonthSheet(spreadsheetId, fecha) {
  const title = getMonthSheetName(fecha);
  await ensureSheet(spreadsheetId, title, HEADER_MUESTRAS);
  return title;
}

async function getAllMonthSheetNames(spreadsheetId) {
  const props = await getSheetProps(spreadsheetId);
  const patron = /^\d{4}-\d{2}$/;
  return props
    .map(p => p.title)
    .filter(t => patron.test(t))
    .sort((a, b) => a.localeCompare(b));
}

/** Filas de datos (sin encabezado). El índice 0 del arreglo devuelto = fila 2 de la hoja. */
async function readAllRows(spreadsheetId, sheetTitle, lastCol) {
  const sheets = await getSheetsClient();
  const col = lastCol || 'J';
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${sheetTitle}'!A2:${col}`
  });
  return res.data.values || [];
}

async function appendRow(spreadsheetId, sheetTitle, row) {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `'${sheetTitle}'!A:A`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });
}

async function appendRows(spreadsheetId, sheetTitle, rows) {
  if (!rows.length) return;
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `'${sheetTitle}'!A:A`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: rows }
  });
}

/** rowNumber1Indexed: número real de fila en la hoja (fila 1 = encabezado). */
async function updateCell(spreadsheetId, sheetTitle, rowNumber1Indexed, colLetter, value) {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${sheetTitle}'!${colLetter}${rowNumber1Indexed}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[value]] }
  });
}

async function deleteRow(spreadsheetId, sheetTitle, rowNumber1Indexed) {
  const sheetId = await getSheetIdByTitle(spreadsheetId, sheetTitle);
  if (sheetId === null) return;
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: rowNumber1Indexed - 1, endIndex: rowNumber1Indexed }
        }
      }]
    }
  });
}

async function deleteRows(spreadsheetId, sheetTitle, rowNumbers1Indexed) {
  if (!rowNumbers1Indexed.length) return;
  const sheetId = await getSheetIdByTitle(spreadsheetId, sheetTitle);
  if (sheetId === null) return;
  const sheets = await getSheetsClient();
  const ordenadas = rowNumbers1Indexed.slice().sort((a, b) => b - a);
  const requests = ordenadas.map(r => ({
    deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: r - 1, endIndex: r } }
  }));
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
}

module.exports = {
  HEADER_MUESTRAS,
  getMonthSheetName,
  getSheetProps,
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
};
