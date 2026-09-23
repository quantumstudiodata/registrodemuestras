// ============================================================
//  Lab Registro de Muestras v3
//  Code.gs
// ============================================================

var SPREADSHEET_ID = '1lbItZAdR7NmAAAB4iykvryPWD4nX13lWwzKEG_ilLLY';
var SHEET_MUESTRAS = 'Muestras';
var TZ = 'America/Monterrey';

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Lab — Registro de Muestras')
    .addMetaTag('viewport','width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_MUESTRAS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MUESTRAS);
    sh.appendRow(['ID','Fecha','Hora','Codigo','Prueba','Analista','Nombre','Ronda','Sucursal','Estado']);
    sh.getRange(1,1,1,10).setFontWeight('bold');
  }
  return { success: true };

}

function getMonthSheetName(fecha) {
  // fecha viene como 'YYYY-MM-DD' -> devuelve 'YYYY-MM'
  return String(fecha).substring(0,7);
}

function getOrCreateMonthSheet(fecha) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var nombre = getMonthSheetName(fecha);
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.appendRow(['ID','Fecha','Hora','Codigo','Prueba','Analista','Nombre','Ronda','Sucursal','Estado']);
    sh.getRange(1,1,1,10).setFontWeight('bold');
  }
  return sh;
}

function getAllMonthSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var todas = ss.getSheets();
  var patron = /^\d{4}-\d{2}$/;
  return todas
    .filter(function(sh){ return patron.test(sh.getName()); })
    .sort(function(a,b){ return a.getName().localeCompare(b.getName()); });
}

function getNow() {
  var OFFSET = -6 * 60 * 60 * 1000;
  var now = new Date(new Date().getTime() + OFFSET);
  return {
    fecha: now.getUTCFullYear()+'-'+String(now.getUTCMonth()+1).padStart(2,'0')+'-'+String(now.getUTCDate()).padStart(2,'0'),
    hora:  String(now.getUTCHours()).padStart(2,'0')+':'+String(now.getUTCMinutes()).padStart(2,'0')+':'+String(now.getUTCSeconds()).padStart(2,'0')
  };
}

function saveMuestra(codigo, pruebas, analista, nombre, ronda) {
  try {
    var t   = getNow();
    var sh  = getOrCreateMonthSheet(t.fecha);
    var ids = [];
    var lista = (Array.isArray(pruebas) && pruebas.length > 0) ? pruebas : [''];
    lista.forEach(function(prueba) {
      var id = Utilities.getUuid();
      sh.appendRow([id, t.fecha, t.hora, String(codigo), String(prueba||''), String(analista||''), String(nombre||''), Number(ronda||1), '', '']);
      ids.push(id);
    });
    return { success: true, ids: ids, fecha: t.fecha, hora: t.hora };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function getMuestrasHoy() {
  try {
    var hojas = getAllMonthSheets();
    var rows = [];

    hojas.forEach(function(sh){
      var data = sh.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        if (!row[0]) continue;

        var fechaRaw = row[1];
        var fecha = (fechaRaw instanceof Date)
          ? Utilities.formatDate(fechaRaw, TZ, 'yyyy-MM-dd')
          : String(fechaRaw).substring(0,10);

        var horaRaw = row[2];
        var hora = (horaRaw instanceof Date)
          ? Utilities.formatDate(horaRaw, TZ, 'HH:mm:ss')
          : String(horaRaw).substring(0,8);

        var estado = String(row[9]||'');
        var prueba = String(row[4]||'');
        rows.push({
          id:        String(row[0]),
          fecha:     fecha,
          hora:      hora,
          codigo:    String(row[3]),
          prueba:    prueba==='RECHAZADA'?'':prueba,
          analista:  String(row[5]||''),
          nombre:    String(row[6]||''),
          ronda:     Number(row[7]||1),
          sucursal:  String(row[8]||''),
          rechazada: prueba==='RECHAZADA',
          manual:    (estado==='manual'||estado==='auto')
        });
      }
    });

    return { success: true, muestras: rows };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function actualizarPruebas(actualizaciones) {
  try {
    var t = getNow();
    var lista = Array.isArray(actualizaciones) ? actualizaciones : [actualizaciones];

    for (var i = 0; i < lista.length; i++) {
      var upd = lista[i];

      if (upd.auto === true || upd.auto === 'true') {
        var fechaAuto = String(upd.fecha || t.fecha);
        var shAuto = getOrCreateMonthSheet(fechaAuto);
        var newIdAuto = Utilities.getUuid();
        var filaAuto = [newIdAuto, fechaAuto, t.hora, String(upd.codigo || ''), String(upd.prueba || ''), String(upd.analista || ''), '', 1, String(upd.sucursal || ''), 'auto'];
        shAuto.appendRow(filaAuto);
        continue;
      }

      if (upd.manual === 'true' || upd.manual === true) {
        var shManual = getOrCreateMonthSheet(t.fecha);
        var newId = Utilities.getUuid();
        var filaManual = [newId, t.fecha, t.hora, String(upd.codigo || ''), String(upd.prueba || ''), String(upd.analista || ''), '', 1, String(upd.sucursal || ''), 'manual'];
        shManual.appendRow(filaManual);
        continue;
      }

      var fechaRegistro = String(upd.fecha || t.fecha);
      var sh = getOrCreateMonthSheet(fechaRegistro);
      var data = sh.getDataRange().getValues();

            if (upd.rechazada === true || upd.rechazada === 'true') {
        for (var r = 1; r < data.length; r++) {
          if (String(data[r][0]) === String(upd.id)) {
            sh.getRange(r + 1, 5).setValue('RECHAZADA');
            if (upd.razon) sh.getRange(r + 1, 7).setValue(String(upd.razon));
            break;
          }
        }
        continue;
      }

      for (var r2 = 1; r2 < data.length; r2++) {
        if (String(data[r2][0]) === String(upd.id)) {
          var pruebaActual = String(data[r2][4] || '');
          if (pruebaActual === '' || pruebaActual === 'undefined') {
            if (upd.prueba !== undefined) sh.getRange(r2 + 1, 5).setValue(upd.prueba);
            if (upd.sucursal) sh.getRange(r2 + 1, 9).setValue(upd.sucursal);
          }
          break;
        }
      }
    }

    sincronizarInventarioHoy();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
function deleteMuestra(id, fecha) {
  try {
    var sh   = getOrCreateMonthSheet(String(fecha||getNow().fecha));
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(id)) {
        sh.deleteRow(r+1);
        return { success: true };
      }
    }
    return { success: false, error: 'No encontrada' };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function guardarManual(payload) {
  try {
    var sh = getOrCreateMonthSheet(String(payload.fecha || getNow().fecha));
    var id = Utilities.getUuid();
    sh.appendRow([
      id,
      String(payload.fecha || ''),
      String(payload.hora || ''),
      String(payload.codigo || ''),
      String(payload.prueba || ''),
      String(payload.analista || ''),
      '',
      1,
      String(payload.sucursal || ''),
      'manual'
    ]);
    sincronizarInventarioHoy();
    return { success: true, id: id };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteByCode(codigo) {
  try {
    var t    = getNow();
    var sh   = getOrCreateMonthSheet(t.fecha);
    var data = sh.getDataRange().getValues();
    var deleted = 0;
    for (var r = data.length - 1; r >= 1; r--) {
      var fechaRaw = data[r][1];
      var fecha = (fechaRaw instanceof Date)
        ? Utilities.formatDate(fechaRaw, TZ, 'yyyy-MM-dd')
        : String(fechaRaw).substring(0,10);
      if (String(data[r][3]) === String(codigo) && fecha === t.fecha) {
        sh.deleteRow(r+1);
        deleted++;
      }
    }
    return { success: true, deleted: deleted };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function migrarAMensual() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var shVieja = ss.getSheetByName(SHEET_MUESTRAS);
  if (!shVieja) {
    return 'No se encontró la pestaña "Muestras" — nada que migrar.';
  }

  var data = shVieja.getDataRange().getValues();
  var migrados = 0;
  var porMes = {}; // agrupamos primero para escribir en bloque (más rápido)

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row[0]) continue; // fila vacía

    var fechaRaw = row[1];
    var fecha = (fechaRaw instanceof Date)
      ? Utilities.formatDate(fechaRaw, TZ, 'yyyy-MM-dd')
      : String(fechaRaw).substring(0,10);

    var mes = getMonthSheetName(fecha);
    if (!porMes[mes]) porMes[mes] = [];
    porMes[mes].push(row);
    migrados++;
  }

  Object.keys(porMes).forEach(function(mes){
    var sh = getOrCreateMonthSheet(porMes[mes][0][1] instanceof Date
      ? Utilities.formatDate(porMes[mes][0][1], TZ, 'yyyy-MM-dd')
      : mes+'-01');
    var filas = porMes[mes];
    // escribir todas las filas de ese mes de un solo golpe
    sh.getRange(sh.getLastRow()+1, 1, filas.length, 10).setValues(filas);
  });

  return 'Migración completa: ' + migrados + ' registros repartidos en ' + Object.keys(porMes).length + ' pestañas de mes. La pestaña "Muestras" original NO se tocó (queda como respaldo).';
}

function getMuestrasMes(mes) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(mes);
    var rows = [];
    if (sh) {
      var data = sh.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        if (!row[0]) continue;
        var fechaRaw = row[1];
        var fecha = (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, TZ, 'yyyy-MM-dd') : String(fechaRaw).substring(0,10);
        var horaRaw = row[2];
        var hora = (horaRaw instanceof Date) ? Utilities.formatDate(horaRaw, TZ, 'HH:mm:ss') : String(horaRaw).substring(0,8);
        var estado = String(row[9]||'');
        var prueba = String(row[4]||'');
        rows.push({
          id: String(row[0]), fecha: fecha, hora: hora, codigo: String(row[3]),
          prueba: prueba==='RECHAZADA'?'':prueba, analista: String(row[5]||''),
          nombre: String(row[6]||''), ronda: Number(row[7]||1), sucursal: String(row[8]||''),
          rechazada: prueba==='RECHAZADA', manual: (estado==='manual'||estado==='auto')
        });
      }
    }
    return { success: true, muestras: rows };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function deleteMuestraGrupo(codigo, fecha) {
  try {
    var sh = getOrCreateMonthSheet(String(fecha||getNow().fecha));
    var data = sh.getDataRange().getValues();
    var eliminadas = 0;
    for (var r = data.length - 1; r >= 1; r--) {
      var fechaRaw = data[r][1];
      var fechaFila = (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, TZ, 'yyyy-MM-dd') : String(fechaRaw).substring(0,10);
      if (String(data[r][3]) === String(codigo) && fechaFila === String(fecha)) {
        sh.deleteRow(r+1);
        eliminadas++;
      }
    }
    return { success: true, eliminadas: eliminadas };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function deleteFilaPorId(id, fecha) {
  try {
    var sh = getOrCreateMonthSheet(String(fecha||getNow().fecha));
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(id)) {
        sh.deleteRow(r+1);
        return { success: true };
      }
    }
    return { success: false, error: 'No encontrada' };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function editarPruebaFila(id, fecha, nuevaPrueba) {
  try {
    var sh = getOrCreateMonthSheet(String(fecha||getNow().fecha));
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(id)) {
        sh.getRange(r+1,5).setValue(String(nuevaPrueba||''));
        return { success: true };
      }
    }
    return { success: false, error: 'No encontrada' };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function renombrarCodigo(codigoViejo, codigoNuevo, fecha, sucursalNueva) {
  try {
    var sh = getOrCreateMonthSheet(String(fecha||getNow().fecha));
    var data = sh.getDataRange().getValues();
    var actualizadas = 0;
    for (var r = 1; r < data.length; r++) {
      var fechaRaw = data[r][1];
      var fechaFila = (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, TZ, 'yyyy-MM-dd') : String(fechaRaw).substring(0,10);
      if (String(data[r][3]) === String(codigoViejo) && fechaFila === String(fecha)) {
        sh.getRange(r+1,4).setValue(String(codigoNuevo));
        if (sucursalNueva) sh.getRange(r+1,9).setValue(String(sucursalNueva));
        actualizadas++;
      }
    }
    return { success: true, actualizadas: actualizadas };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

/** Prueba manual: ejecútala con ▶️ desde el editor para confirmar que la Biblioteca funciona. */
function probarLibreriaInventario() {
  var resultado = InventarioLib.aplicarReceta('COPROLOGICO', 1);
  Logger.log(JSON.stringify(resultado));
}
var SHEET_PERFILES = 'Perfiles';
var SHEET_MURO = 'Muro';
var CARPETA_FOTOS_ID = null; // se llena solo la primera vez

function getOrCreateCarpetaFotos() {
  if (CARPETA_FOTOS_ID) return DriveApp.getFolderById(CARPETA_FOTOS_ID);
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CARPETA_FOTOS_ID');
  if (id) { CARPETA_FOTOS_ID = id; return DriveApp.getFolderById(id); }
  var folder = DriveApp.createFolder('Fotos de Perfil - Lab');
  props.setProperty('CARPETA_FOTOS_ID', folder.getId());
  CARPETA_FOTOS_ID = folder.getId();
  return folder;
}

function getOrCreatePerfilesSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_PERFILES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PERFILES);
    sh.appendRow(['Usuario','NombreCompleto','FotoURL']);
    sh.getRange(1,1,1,3).setFontWeight('bold');
  }
  return sh;
}

function getOrCreateMuroSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_MURO);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MURO);
    sh.appendRow(['ID','Fecha','Hora','Usuario','Mensaje']);
    sh.getRange(1,1,1,5).setFontWeight('bold');
  }
  return sh;
}

function getPerfil(usuario) {
  try {
    var sh = getOrCreatePerfilesSheet();
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(usuario)) {
        return { success: true, nombreCompleto: String(data[r][1]||''), fotoURL: String(data[r][2]||'') };
      }
    }
    return { success: true, nombreCompleto: '', fotoURL: '' };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function guardarPerfil(usuario, nombreCompleto, fotoBase64) {
  try {
    var sh = getOrCreatePerfilesSheet();
    var data = sh.getDataRange().getValues();
    var fotoURL = '';

    if (fotoBase64) {
      var partes = fotoBase64.split(',');
      var meta = partes[0];
      var contenido = partes[1];
      var mime = meta.match(/data:(.*);base64/)[1];
      var bytes = Utilities.base64Decode(contenido);
      var blob = Utilities.newBlob(bytes, mime, usuario + '_foto');
      var carpeta = getOrCreateCarpetaFotos();
      var archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fotoURL = 'https://drive.google.com/thumbnail?id=' + archivo.getId() + '&sz=w200';
    }

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(usuario)) {
        sh.getRange(r+1,2).setValue(nombreCompleto||'');
        if (fotoURL) sh.getRange(r+1,3).setValue(fotoURL);
        return { success: true, fotoURL: fotoURL || String(data[r][2]||'') };
      }
    }
    sh.appendRow([usuario, nombreCompleto||'', fotoURL]);
    return { success: true, fotoURL: fotoURL };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function getMuro() {
  try {
    var sh = getOrCreateMuroSheet();
    var data = sh.getDataRange().getValues();
    var mensajes = [];
    for (var r = 1; r < data.length; r++) {
      if (!data[r][0]) continue;
      mensajes.push({
        id: String(data[r][0]),
        fecha: String(data[r][1]),
        hora: String(data[r][2]),
        usuario: String(data[r][3]),
        mensaje: String(data[r][4])
      });
    }
    return { success: true, mensajes: mensajes };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function publicarMensaje(usuario, mensaje) {
  try {
    var sh = getOrCreateMuroSheet();
    var t = getNow();
    var id = Utilities.getUuid();
    sh.appendRow([id, t.fecha, t.hora, usuario, mensaje]);
    return { success: true, id: id, fecha: t.fecha, hora: t.hora };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function eliminarMensajeMuro(id) {
  try {
    var sh = getOrCreateMuroSheet();
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(id)) {
        sh.deleteRow(r+1);
        return { success: true };
      }
    }
    return { success: false, error: 'No encontrado' };
  } catch(err) {
    return { success: false, error: err.message };
  }
}
function getMuestrasMeses(meses) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var rows = [];
    meses.forEach(function(mes){
      var sh = ss.getSheetByName(mes);
      if (!sh) return;
      var data = sh.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        if (!row[0]) continue;
        var fechaRaw = row[1];
        var fecha = (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, TZ, 'yyyy-MM-dd') : String(fechaRaw).substring(0,10);
        var horaRaw = row[2];
        var hora = (horaRaw instanceof Date) ? Utilities.formatDate(horaRaw, TZ, 'HH:mm:ss') : String(horaRaw).substring(0,8);
        var estado = String(row[9]||'');
        var prueba = String(row[4]||'');
        rows.push({
          id: String(row[0]), fecha: fecha, hora: hora, codigo: String(row[3]),
          prueba: prueba==='RECHAZADA'?'':prueba, analista: String(row[5]||''),
          nombre: String(row[6]||''), ronda: Number(row[7]||1), sucursal: String(row[8]||''),
          rechazada: prueba==='RECHAZADA', manual: (estado==='manual'||estado==='auto')
        });
      }
    });
    return { success: true, muestras: rows };
  } catch(err) {
    return { success: false, error: err.message };
  }
}
var PRUEBAS_ESTADISTICAS_ = {
  'COPROLOGICO': ['COPROLOGICO','CPS','SOH','AMF'],
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

function extraerCodigoPaciente_(cod) {
  var s = String(cod || '').split(/[-']/)[0];
  var r = s.length > 2 ? s.substring(2) : s;
  return r.replace(/\.0$/, '');
}

function getStatsHoyServidor_() {
  var t = getNow();
  var sh = getOrCreateMonthSheet(t.fecha);
  var data = sh.getDataRange().getValues();
  var counts = {};
  var vistos = {};

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row[0]) continue;

    var fechaRaw = row[1];
    var fecha = (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, TZ, 'yyyy-MM-dd') : String(fechaRaw).substring(0, 10);
    if (fecha !== t.fecha) continue;

    var prueba = String(row[4] || '');
    if (!prueba || prueba === 'RECHAZADA') continue;

    var cats = PRUEBAS_ESTADISTICAS_[prueba];
    if (!cats) continue;

    var codigoPac = extraerCodigoPaciente_(row[3]);
    cats.forEach(function (cat) {
      var clave = codigoPac + '_' + prueba + '_' + cat;
      if (!vistos[clave]) {
        vistos[clave] = true;
        counts[cat] = (counts[cat] || 0) + 1;
      }
    });
  }

  return counts;
}

function sincronizarInventarioHoy() {
  try {
    var stats = getStatsHoyServidor_();
    return InventarioLib.sincronizarTotalesDelDia(stats);
  } catch (err) {
    Logger.log('No se pudo sincronizar el inventario: ' + err.message);
    return null;
  }
}

function limpiarDuplicadosARPHH() {
  var hojas = getAllMonthSheets();
  var totalEliminados = 0;
  var detalle = [];

  hojas.forEach(function(sh){
    var data = sh.getDataRange().getValues();
    var codigosConCopro = {};
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][4]) === 'COPROLOGICO') {
        codigosConCopro[String(data[r][3])] = true;
      }
    }

    var filasABorrar = [];
    for (var r2 = 1; r2 < data.length; r2++) {
      var prueba = String(data[r2][4]);
      var estado = String(data[r2][9]);
      var codigo = String(data[r2][3]);
      if ((prueba === 'AR' || prueba === 'PHH') && estado === 'auto' && codigosConCopro[codigo]) {
        filasABorrar.push(r2 + 1); // +1 porque las filas del sheet empiezan en 1
      }
    }

    // Borrar de abajo hacia arriba para no desfasar los índices
    filasABorrar.sort(function(a,b){ return b - a; });
    filasABorrar.forEach(function(fila){
      sh.deleteRow(fila);
      totalEliminados++;
    });

    if (filasABorrar.length > 0) {
      detalle.push(sh.getName() + ': ' + filasABorrar.length + ' filas eliminadas');
    }
  });

  Logger.log('Total eliminado: ' + totalEliminados);
  Logger.log(detalle.join('\n'));
  return totalEliminados + ' filas eliminadas. Detalle: ' + detalle.join(' | ');
}
