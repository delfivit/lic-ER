/*************************************************************************
 *  LICITACIONES ENTRE RÍOS — Granss SRL
 *  Script de la PLANILLA (no scrapea: eso lo hace GitHub todos los días)
 *
 *  Qué hace:
 *   · Baja el listado que arma GitHub y lo vuelca en la hoja "Licitaciones"
 *   · NUNCA pisa tus columnas de seguimiento (Estado propio, precio, posición,
 *     notas). Cada licitación se reconoce por un ID interno, así que podés
 *     escribir tranquila que no se te borra nada.
 *   · Arma la hoja "Resumen" con lo que vence primero y cómo venimos.
 *
 *  Instalación (una vez):
 *   1) Extensiones > Apps Script, pegar TODO esto reemplazando lo que haya
 *   2) Ejecutar  instalarTodo
 *************************************************************************/

var CFG = {
  CSV: 'https://raw.githubusercontent.com/delfivit/lic-ER/main/licitaciones.csv',
  HOJA: 'Licitaciones',
  RESUMEN: 'Resumen',
  HORA: 9                      // hora a la que se actualiza sola (GitHub corre 8:00)
};

// Columnas que llena el sistema. No las edites: se reescriben en cada actualización.
var AUTO = ['Apertura','Venta pliego hasta','Estado','Rubro','Organismo','Objeto'];
// Columnas tuyas. El script jamás las toca.
var MIAS = ['Seguimiento','Precio ofertado','Posición','Notas'];
// El resto de los datos automáticos.
var AUTO2 = ['N°','Tipo','Valor pliego','Detalle apertura','Link','Pliego PDF','Fuente','Detectada','ID'];

var COLS = AUTO.concat(MIAS).concat(AUTO2);
var ESTADOS = ['', 'A revisar', 'Presentada', 'Descartada', 'Ganada', 'Perdida'];

// del nombre de la columna del CSV al de la planilla
var DESDE_CSV = {
  'Apertura':'Apertura', 'Venta hasta':'Venta pliego hasta', 'Estado':'Estado',
  'Rubro':'Rubro', 'Organismo':'Organismo', 'Objeto':'Objeto', 'N°':'N°',
  'Tipo':'Tipo', 'Valor pliego':'Valor pliego', 'Detalle apertura':'Detalle apertura',
  'Link':'Link', 'Pliego PDF':'Pliego PDF', 'Fuente':'Fuente',
  'Detectada':'Detectada', 'ID':'ID'
};

// ====================== LO PRINCIPAL ==================================
function actualizar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var filas = bajarCsv_();
  if (!filas.length) {
    ss.toast('No pude bajar el listado. ¿Hay internet? Probá de nuevo en un rato.', 'Licitaciones', 10);
    return 0;
  }
  var hoja = prepararHoja_(ss);

  // 1) Guardar lo que vos escribiste, indexado por ID
  var mio = {};
  var ultima = hoja.getLastRow();
  if (ultima > 1) {
    var previo = hoja.getRange(2, 1, ultima - 1, COLS.length).getValues();
    for (var i = 0; i < previo.length; i++) {
      var id = (previo[i][COLS.indexOf('ID')] || '').toString().trim();
      if (!id) continue;
      var guardado = {};
      for (var m = 0; m < MIAS.length; m++) guardado[MIAS[m]] = previo[i][COLS.indexOf(MIAS[m])];
      mio[id] = guardado;
    }
  }

  // 2) Armar las filas nuevas, devolviendo a cada una lo tuyo
  var salida = [], conSeguimiento = 0;
  for (var f = 0; f < filas.length; f++) {
    var reg = filas[f];
    var id = (reg['ID'] || '').toString().trim();
    var fila = new Array(COLS.length).fill('');
    for (var c in DESDE_CSV) {
      var destino = COLS.indexOf(DESDE_CSV[c]);
      if (destino >= 0) fila[destino] = reg[c] || '';
    }
    var g = mio[id];
    if (g) {
      for (var k = 0; k < MIAS.length; k++) fila[COLS.indexOf(MIAS[k])] = g[MIAS[k]] || '';
      if ((g['Seguimiento'] || '').toString().trim()) conSeguimiento++;
    }
    salida.push(fila);
  }

  // 3) Ordenar: primero lo que abre antes; lo que no tiene fecha, al final
  salida.sort(function (a, b) {
    var fa = aFecha_(a[COLS.indexOf('Apertura')]), fb = aFecha_(b[COLS.indexOf('Apertura')]);
    if (fa && fb) return fa - fb;
    if (fa) return -1;
    if (fb) return 1;
    return 0;
  });

  // 4) Escribir
  if (ultima > 1) hoja.getRange(2, 1, ultima - 1, COLS.length).clearContent();
  if (salida.length) {
    forzarTexto_(hoja, salida.length);              // que no convierta "08/2026" en fecha
    hoja.getRange(2, 1, salida.length, COLS.length).setValues(salida);
  }
  pintar_(hoja, salida.length);
  armarResumen_(ss, salida);

  ss.toast(salida.length + ' licitaciones · ' + conSeguimiento + ' con seguimiento tuyo',
           'Actualizado', 8);
  return salida.length;
}

// ====================== BAJAR EL CSV ==================================
function bajarCsv_() {
  try {
    var r = UrlFetchApp.fetch(CFG.CSV + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return [];
    var tabla = Utilities.parseCsv(r.getContentText());
    if (tabla.length < 2) return [];
    var cab = tabla[0], out = [];
    for (var i = 1; i < tabla.length; i++) {
      var o = {};
      for (var c = 0; c < cab.length; c++) o[cab[c]] = tabla[i][c];
      out.push(o);
    }
    return out;
  } catch (e) {
    Logger.log('Error bajando el CSV: ' + e);
    return [];
  }
}

// ====================== LA HOJA =======================================
function prepararHoja_(ss) {
  var hoja = ss.getSheetByName(CFG.HOJA);
  if (!hoja) hoja = ss.insertSheet(CFG.HOJA, 0);

  var cab = hoja.getRange(1, 1, 1, COLS.length).getValues()[0];
  var igual = true;
  for (var i = 0; i < COLS.length; i++) if (cab[i] !== COLS[i]) { igual = false; break; }
  if (!igual) {
    hoja.getRange(1, 1, 1, COLS.length).setValues([COLS]);
  }

  // encabezado: azul lo automático, verde lo tuyo (para que se note qué podés editar)
  for (var c = 0; c < COLS.length; c++) {
    var esMia = MIAS.indexOf(COLS[c]) >= 0;
    hoja.getRange(1, c + 1)
        .setFontWeight('bold').setFontColor('#ffffff')
        .setBackground(esMia ? '#1E7B4F' : '#1F3A5F')
        .setVerticalAlignment('middle').setWrap(true);
  }
  hoja.setFrozenRows(1);
  hoja.setFrozenColumns(1);

  var ancho = {'Apertura':90,'Venta pliego hasta':105,'Estado':95,'Rubro':115,'Organismo':230,
               'Objeto':420,'Seguimiento':115,'Precio ofertado':120,'Posición':80,'Notas':220,
               'N°':85,'Tipo':130,'Valor pliego':190,'Detalle apertura':240,'Link':110,
               'Pliego PDF':90,'Fuente':170,'Detectada':90,'ID':40};
  for (var a in ancho) if (COLS.indexOf(a) >= 0) hoja.setColumnWidth(COLS.indexOf(a) + 1, ancho[a]);
  hoja.hideColumns(COLS.indexOf('ID') + 1);
  return hoja;
}

function forzarTexto_(hoja, n) {
  // Sin esto Sheets convierte "08/2026" (un N° de licitación) en una fecha
  // y arruina el dato. Pasó de verdad, por eso está.
  var texto = ['N°','Tipo','Objeto','Detalle apertura','Valor pliego','Fuente','ID',
               'Organismo','Rubro','Estado','Apertura','Venta pliego hasta','Detectada'];
  for (var i = 0; i < texto.length; i++) {
    var c = COLS.indexOf(texto[i]);
    if (c >= 0) hoja.getRange(2, c + 1, n, 1).setNumberFormat('@');
  }
  hoja.getRange(2, COLS.indexOf('Precio ofertado') + 1, n, 1).setNumberFormat('$#,##0.00');
}

function pintar_(hoja, n) {
  if (!n) return;
  var rango = hoja.getRange(2, 1, n, COLS.length);
  rango.setFontColor('#000000').setFontWeight('normal').setBackground('#ffffff')
       .setVerticalAlignment('top').setWrap(false);
  hoja.getRange(2, COLS.indexOf('Objeto') + 1, n, 1).setWrap(true);
  hoja.getRange(2, COLS.indexOf('Notas') + 1, n, 1).setWrap(true);

  var cRubro = COLS.indexOf('Rubro') + 1, cEstado = COLS.indexOf('Estado') + 1;
  var cSeg = COLS.indexOf('Seguimiento') + 1, cApert = COLS.indexOf('Apertura') + 1;
  var rubros = hoja.getRange(2, cRubro, n, 1).getValues();
  var estados = hoja.getRange(2, cEstado, n, 1).getValues();
  var segs = hoja.getRange(2, cSeg, n, 1).getValues();
  var aperturas = hoja.getRange(2, cApert, n, 1).getValues();

  var colorRubro = {'Arquitectura':'#D9E1F2','Infraestructura':'#E2EFDA',
                    'Áridos':'#FCE4D6','Vehículos':'#FFF2CC'};
  var colorSeg = {'Presentada':'#CFE2F3','Ganada':'#B7E1CD','Perdida':'#F4CCCC',
                  'Descartada':'#EFEFEF','A revisar':'#FFF2CC'};

  var bgR = [], bgE = [], bgS = [], fwE = [], bgA = [];
  var hoy = new Date(); hoy.setHours(0,0,0,0);
  for (var i = 0; i < n; i++) {
    bgR.push([colorRubro[rubros[i][0]] || '#ffffff']);
    var est = estados[i][0];
    bgE.push([est === 'Vigente' ? '#D9EAD3' : (est === 'Vencida' ? '#F4CCCC' : '#FFF2CC')]);
    fwE.push([est === 'Vigente' ? 'bold' : 'normal']);
    bgS.push([colorSeg[segs[i][0]] || '#ffffff']);
    // apertura en rojo si es dentro de los próximos 7 días
    var f = aFecha_(aperturas[i][0]);
    var dias = f ? (f - hoy) / 86400000 : null;
    bgA.push([(dias !== null && dias >= 0 && dias <= 7) ? '#FCE4D6' : '#ffffff']);
  }
  hoja.getRange(2, cRubro, n, 1).setBackgrounds(bgR).setFontWeight('bold');
  hoja.getRange(2, cEstado, n, 1).setBackgrounds(bgE).setFontWeights(fwE);
  hoja.getRange(2, cSeg, n, 1).setBackgrounds(bgS);
  hoja.getRange(2, cApert, n, 1).setBackgrounds(bgA).setFontWeight('bold');

  // desplegable en Seguimiento
  var val = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true)
            .setAllowInvalid(false)
            .setHelpText('Elegí: A revisar, Presentada, Descartada, Ganada o Perdida').build();
  hoja.getRange(2, cSeg, Math.max(n, 200), 1).setDataValidation(val);

  // links clickeables
  linkear_(hoja, n, 'Link', 'ver');
  linkear_(hoja, n, 'Pliego PDF', 'PDF');

  hoja.getRange(1, 1, n + 1, COLS.length).createFilter();
}

function linkear_(hoja, n, columna, texto) {
  var c = COLS.indexOf(columna) + 1;
  if (c <= 0) return;
  var vals = hoja.getRange(2, c, n, 1).getValues();
  var out = [];
  for (var i = 0; i < n; i++) {
    var u = (vals[i][0] || '').toString();
    out.push([/^https?:\/\//.test(u) ? '=HYPERLINK("' + u.replace(/"/g, '%22') + '";"' + texto + '")' : '']);
  }
  hoja.getRange(2, c, n, 1).setFormulas(out);
}

function aFecha_(v) {
  if (v instanceof Date) return v;
  var m = (v || '').toString().match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return null;
  var d = new Date(+m[3], +m[2] - 1, +m[1]);
  d.setHours(0, 0, 0, 0);
  return d;
}

// ====================== RESUMEN =======================================
function armarResumen_(ss, filas) {
  var sh = ss.getSheetByName(CFG.RESUMEN) || ss.insertSheet(CFG.RESUMEN, 1);
  sh.clear();

  var iEst = COLS.indexOf('Estado'), iRub = COLS.indexOf('Rubro'),
      iAp = COLS.indexOf('Apertura'), iSeg = COLS.indexOf('Seguimiento'),
      iOrg = COLS.indexOf('Organismo'), iObj = COLS.indexOf('Objeto'),
      iVenta = COLS.indexOf('Venta pliego hasta'), iPrec = COLS.indexOf('Precio ofertado');

  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  var vig = [], porRubro = {}, porSeg = {}, prox = [], conFechaPliego = 0;
  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    var seg = (f[iSeg] || '').toString().trim();
    if (seg) porSeg[seg] = (porSeg[seg] || 0) + 1;
    if (f[iEst] !== 'Vigente') continue;
    vig.push(f);
    porRubro[f[iRub]] = (porRubro[f[iRub]] || 0) + 1;
    var v = (f[iVenta] || '').toString();
    if (v && v !== 'Ver pliego') conFechaPliego++;
    var fa = aFecha_(f[iAp]);
    if (fa) {
      var dias = Math.round((fa - hoy) / 86400000);
      if (dias >= 0 && dias <= 15) prox.push([dias, f]);
    }
  }
  prox.sort(function (a, b) { return a[0] - b[0]; });

  var fila = 1;
  function titulo(txt, color) {
    sh.getRange(fila, 1, 1, 6).merge().setValue(txt)
      .setFontSize(13).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground(color || '#1F3A5F').setVerticalAlignment('middle');
    sh.setRowHeight(fila, 26); fila += 2;
  }
  function encabezado(arr) {
    sh.getRange(fila, 1, 1, arr.length).setValues([arr])
      .setFontWeight('bold').setBackground('#E8EAED');
    fila++;
  }

  titulo('LICITACIONES ENTRE RÍOS — Granss SRL');
  sh.getRange(fila - 1, 1).setValue(
    'Actualizado: ' + Utilities.formatDate(new Date(), 'America/Argentina/Cordoba', 'dd/MM/yyyy HH:mm') +
    '   ·   Se actualiza solo todos los días').setFontColor('#666666');
  fila += 1;

  // --- números grandes
  encabezado(['VIGENTES', 'Con fecha de apertura', 'Con fecha de venta de pliego', 'En seguimiento', '', '']);
  sh.getRange(fila, 1, 1, 4).setValues([[
    vig.length,
    vig.filter(function (f) { return aFecha_(f[iAp]); }).length,
    conFechaPliego,
    Object.keys(porSeg).reduce(function (t, k) { return t + porSeg[k]; }, 0)
  ]]).setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
  fila += 3;

  // --- por rubro
  titulo('VIGENTES POR RUBRO', '#3C6E9F');
  encabezado(['Rubro', 'Cantidad']);
  var colorRubro = {'Arquitectura':'#D9E1F2','Infraestructura':'#E2EFDA','Áridos':'#FCE4D6','Vehículos':'#FFF2CC'};
  Object.keys(porRubro).sort(function (a, b) { return porRubro[b] - porRubro[a]; })
    .forEach(function (r) {
      sh.getRange(fila, 1, 1, 2).setValues([[r, porRubro[r]]]);
      sh.getRange(fila, 1).setBackground(colorRubro[r] || '#ffffff').setFontWeight('bold');
      fila++;
    });
  fila += 1;

  // --- cómo venimos
  titulo('CÓMO VENIMOS', '#1E7B4F');
  encabezado(['Estado', 'Cantidad']);
  var colorSeg = {'Presentada':'#CFE2F3','Ganada':'#B7E1CD','Perdida':'#F4CCCC',
                  'Descartada':'#EFEFEF','A revisar':'#FFF2CC'};
  var hayAlguno = false;
  ['A revisar', 'Presentada', 'Ganada', 'Perdida', 'Descartada'].forEach(function (e) {
    if (!porSeg[e]) return;
    hayAlguno = true;
    sh.getRange(fila, 1, 1, 2).setValues([[e, porSeg[e]]]);
    sh.getRange(fila, 1).setBackground(colorSeg[e]).setFontWeight('bold');
    fila++;
  });
  if (!hayAlguno) {
    sh.getRange(fila, 1, 1, 4).merge()
      .setValue('Todavía no marcaste ninguna. Usá la columna "Seguimiento" (verde) en la hoja Licitaciones.')
      .setFontColor('#888888').setFontStyle('italic');
    fila++;
  }
  var ganadas = porSeg['Ganada'] || 0, perdidas = porSeg['Perdida'] || 0;
  if (ganadas + perdidas > 0) {
    sh.getRange(fila, 1, 1, 2).setValues([['Efectividad', Math.round(ganadas * 100 / (ganadas + perdidas)) + '%']]);
    sh.getRange(fila, 1, 1, 2).setFontWeight('bold');
    fila++;
  }
  fila += 1;

  // --- las que vencen primero
  titulo('ABREN EN LOS PRÓXIMOS 15 DÍAS', '#B45309');
  if (!prox.length) {
    sh.getRange(fila, 1, 1, 6).merge().setValue('No hay aperturas en los próximos 15 días.')
      .setFontColor('#888888').setFontStyle('italic');
    fila++;
  } else {
    encabezado(['Faltan', 'Apertura', 'Venta pliego', 'Rubro', 'Organismo', 'Objeto']);
    prox.forEach(function (p) {
      var d = p[0], f = p[1];
      sh.getRange(fila, 1, 1, 6).setValues([[
        d === 0 ? '¡HOY!' : (d === 1 ? 'mañana' : d + ' días'),
        f[iAp], f[iVenta], f[iRub], (f[iOrg] || '').toString().slice(0, 45),
        (f[iObj] || '').toString().slice(0, 110)
      ]]);
      if (d <= 3) sh.getRange(fila, 1, 1, 6).setBackground('#FCE4D6');
      sh.getRange(fila, 1).setFontWeight('bold');
      fila++;
    });
  }

  sh.setColumnWidth(1, 95); sh.setColumnWidth(2, 100); sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 260); sh.setColumnWidth(6, 460);
  sh.getRange(1, 1, fila, 6).setVerticalAlignment('middle');
  sh.setFrozenRows(3);
}

// ====================== ORDENAR LA PLANILLA ===========================
/**
 * Borra las hojas viejas que quedaron de versiones anteriores.
 * Pregunta antes de borrar nada.
 */
function limpiarHojasViejas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quedan = [CFG.HOJA, CFG.RESUMEN, 'Mails'];
  var aBorrar = ss.getSheets().filter(function (h) { return quedan.indexOf(h.getName()) < 0; });
  if (!aBorrar.length) { ss.toast('No hay hojas viejas para borrar.', 'Limpieza', 6); return; }

  var nombres = aBorrar.map(function (h) { return '• ' + h.getName(); }).join('\n');
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('Borrar hojas viejas',
    'Se van a borrar estas hojas:\n\n' + nombres +
    '\n\nSe conservan: ' + quedan.join(', ') +
    '\n\n¿Confirmás? (esto no se puede deshacer)', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) { ss.toast('No se borró nada.', 'Limpieza', 5); return; }

  aBorrar.forEach(function (h) { ss.deleteSheet(h); });
  ss.toast(aBorrar.length + ' hoja(s) borrada(s).', 'Limpieza', 8);
}

// ====================== INSTALACIÓN ===================================
function instalarTodo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // sacar cualquier disparador viejo (incluido el del scraper anterior)
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('actualizar').timeBased().everyDays(1).atHour(CFG.HORA).create();
  var n = actualizar();
  ss.toast('Listo: ' + n + ' licitaciones. Se actualiza sola todos los días a las ' + CFG.HORA + ':00.',
           'Instalado', 12);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏗️ Licitaciones')
    .addItem('Actualizar ahora', 'actualizar')
    .addSeparator()
    .addItem('Borrar hojas viejas', 'limpiarHojasViejas')
    .addItem('Instalar / reinstalar automático', 'instalarTodo')
    .addToUi();
}
