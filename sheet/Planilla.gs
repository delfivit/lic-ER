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
  // Se lee por raw.githubusercontent y la API queda de respaldo. La API es más
  // fresca (raw cachea unos minutos) PERO sólo admite 60 pedidos por hora y por
  // IP, y Apps Script sale por IPs de Google compartidas con miles de usuarios:
  // el cupo se agota y devuelve vacío. Como esto corre una vez por día, esos
  // minutos de caché no molestan.
  CSV: 'https://raw.githubusercontent.com/delfivit/lic-ER/main/licitaciones.csv',
  CSV_RESPALDO: 'https://api.github.com/repos/delfivit/lic-ER/contents/licitaciones.csv',
  DIAG: 'https://raw.githubusercontent.com/delfivit/lic-ER/main/diagnostico.csv',
  DIAG_RESPALDO: 'https://api.github.com/repos/delfivit/lic-ER/contents/diagnostico.csv',
  FUENTES: 'Fuentes',
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
  var estadoFuentes = actualizarFuentes_(ss);
  armarResumen_(ss, salida, estadoFuentes);

  var aviso = salida.length + ' licitaciones · ' + conSeguimiento + ' con seguimiento tuyo';
  if (estadoFuentes.conError) {
    aviso += '  ⚠️ ' + estadoFuentes.conError + ' fuente(s) con problemas: mirá la hoja Fuentes';
  }
  ss.toast(aviso, 'Actualizado', 12);
  return salida.length;
}

// ====================== BAJAR EL CSV ==================================
function bajarCsv_() {
  var texto = pedir_(CFG.CSV + '?t=' + Date.now(), {});
  if (!texto) texto = pedir_(CFG.CSV_RESPALDO, { 'Accept': 'application/vnd.github.raw' });
  if (!texto) return [];
  try {
    var tabla = Utilities.parseCsv(texto);
    if (tabla.length < 2) return [];
    var cab = tabla[0], out = [];
    for (var i = 1; i < tabla.length; i++) {
      var o = {};
      for (var c = 0; c < cab.length; c++) o[cab[c]] = tabla[i][c];
      out.push(o);
    }
    return out;
  } catch (e) {
    Logger.log('Error leyendo el CSV: ' + e);
    return [];
  }
}

function pedir_(url, headers) {
  try {
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: headers });
    return r.getResponseCode() === 200 ? r.getContentText() : null;
  } catch (e) {
    Logger.log('No pude bajar ' + url + ': ' + e);
    return null;
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

  // createFilter() explota si la hoja ya tiene uno, y esto corre en cada
  // actualización: hay que sacar el anterior primero.
  var filtro = hoja.getFilter();
  if (filtro) filtro.remove();
  hoja.getRange(1, 1, n + 1, COLS.length).createFilter();
}

function linkear_(hoja, n, columna, texto) {
  var c = COLS.indexOf(columna) + 1;
  if (c <= 0) return;
  var vals = hoja.getRange(2, c, n, 1).getValues();
  var out = [];
  for (var i = 0; i < n; i++) {
    var u = (vals[i][0] || '').toString();
    if (!/^https?:\/\//.test(u)) { out.push(['']); continue; }
    // Los avisos del Boletín traen #page=N. El boletín tiene ~100 páginas, así
    // que además de abrir ahí mostramos el número por si el visor no salta.
    var pag = u.match(/#page=(\d+)/);
    var rotulo = pag ? (texto + ' pág. ' + pag[1]) : texto;
    out.push(['=HYPERLINK("' + u.replace(/"/g, '%22') + '";"' + rotulo + '")']);
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
function armarResumen_(ss, filas, estadoFuentes) {
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
  if (estadoFuentes && estadoFuentes.conError) {
    sh.getRange(fila - 1, 1, 1, 6).merge()
      .setValue('⚠️  ' + estadoFuentes.conError + ' fuente(s) NO se pudieron revisar. ' +
                'Mirá la hoja "Fuentes" (filas rojas o naranjas) y entrá a esas webs a mano ' +
                'para no perderte licitaciones.')
      .setBackground('#F4CCCC').setFontWeight('bold').setWrap(true);
    sh.setRowHeight(fila - 1, 34);
    fila += 2;
  }
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

// ====================== HOJA FUENTES ==================================
// Columnas que edita el equipo (verde) y columnas que llena el sistema (azul).
var F_MIAS = ['Activa','Nombre','URL','Parser','Notas'];
var F_AUTO = ['Estado','Encontradas','De tu rubro','Corridas sin traer nada','Detalle','Última corrida'];
var F_COLS = F_MIAS.concat(F_AUTO);
var PARSERS = ['boletin','wpjson','generico','parana','minplan','iapv','enersa'];

function actualizarFuentes_(ss) {
  var diag = bajarDiagnostico_();          // lo que reportó la última corrida
  var sh = ss.getSheetByName(CFG.FUENTES);
  var nueva = !sh;
  if (!sh) sh = ss.insertSheet(CFG.FUENTES);

  // lo que ya está cargado en la hoja (para no pisar lo que escribió el equipo)
  var previas = [], ultima = sh.getLastRow();
  if (!nueva && ultima > 1) {
    var v = sh.getRange(2, 1, ultima - 1, F_COLS.length).getValues();
    for (var i = 0; i < v.length; i++) {
      var url = (v[i][F_COLS.indexOf('URL')] || '').toString().trim();
      if (url) previas.push({
        activa: v[i][F_COLS.indexOf('Activa')], nombre: v[i][F_COLS.indexOf('Nombre')],
        url: url, parser: v[i][F_COLS.indexOf('Parser')], notas: v[i][F_COLS.indexOf('Notas')]
      });
    }
  }

  // unir: manda lo que está en la hoja; el diagnóstico aporta el estado
  var porUrl = {};
  previas.forEach(function (p) { porUrl[norUrl_(p.url)] = p; });
  var filas = [], vistas = {};

  diag.forEach(function (d) {
    var k = norUrl_(d['URL']);
    var p = porUrl[k];
    vistas[k] = true;
    filas.push([
      p ? p.activa : (d['Estado'] === 'Apagada' ? 'NO' : 'SI'),
      p && p.nombre ? p.nombre : d['Fuente'],
      d['URL'],
      p && p.parser ? p.parser : d['Parser'],
      p && p.notas ? p.notas : d['Detalle'],
      d['Estado'], d['Encontradas'], d['De tu rubro'],
      d['Corridas sin traer nada'], d['Detalle'], d['Última corrida']
    ]);
  });
  // las que el equipo agregó y todavía no se buscaron
  var sinDiag = (diag.length === 0);
  previas.forEach(function (p) {
    if (vistas[norUrl_(p.url)]) return;
    filas.push([p.activa, p.nombre, p.url, p.parser, p.notas,
                sinDiag ? 'Sin datos' : 'Se busca mañana', '', '', '',
                sinDiag ? 'No se pudo leer el estado de las fuentes. Probá "Actualizar ahora" en un rato.'
                        : 'Agregada por el equipo — todavía no se probó', '']);
  });

  // Limpiar TODO antes de reescribir. Si sólo se borra el bloque de datos, el
  // texto de ayuda del pie queda pegado más abajo y se va duplicando en cada
  // corrida. Lo que escribió el equipo ya está a salvo en `previas`.
  sh.clear();
  sh.getRange(1, 1, 1, F_COLS.length).setValues([F_COLS]);
  for (var c = 0; c < F_COLS.length; c++) {
    sh.getRange(1, c + 1).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground(F_MIAS.indexOf(F_COLS[c]) >= 0 ? '#1E7B4F' : '#1F3A5F').setWrap(true);
  }
  var conError = 0;
  if (filas.length) {
    sh.getRange(2, 1, filas.length, F_COLS.length).setValues(filas)
      .setFontColor('#000000').setBackground('#ffffff').setVerticalAlignment('top').setWrap(true);
    var fondos = [];
    for (var i = 0; i < filas.length; i++) {
      var est = (filas[i][F_COLS.indexOf('Estado')] || '').toString();
      var sinNada = Number(filas[i][F_COLS.indexOf('Corridas sin traer nada')]) || 0;
      var color = '#ffffff';
      if (est === 'ERROR') { color = '#F4CCCC'; conError++; }
      else if (sinNada >= 3) { color = '#FCE4D6'; conError++; }
      else if (est === 'Apagada') color = '#EFEFEF';
      else if (est === 'Se busca mañana') color = '#D9E1F2';
      else if (est === 'OK') color = '#D9EAD3';
      fondos.push(new Array(F_COLS.length).fill(color));
    }
    sh.getRange(2, 1, filas.length, F_COLS.length).setBackgrounds(fondos);

    var vSi = SpreadsheetApp.newDataValidation().requireValueInList(['SI','NO'], true)
              .setAllowInvalid(false).build();
    sh.getRange(2, 1, Math.max(filas.length, 60), 1).setDataValidation(vSi);
    var vP = SpreadsheetApp.newDataValidation().requireValueInList(PARSERS, true)
             .setAllowInvalid(true)
             .setHelpText('wpjson = sitios WordPress (probá este primero) · generico = cualquier otro').build();
    sh.getRange(2, 4, Math.max(filas.length, 60), 1).setDataValidation(vP);
  }

  // instrucciones al pie
  var f = filas.length + 3;
  sh.getRange(f, 1, 1, 6).merge().setValue('CÓMO AGREGAR UN SITIO NUEVO')
    .setFontWeight('bold').setBackground('#1E7B4F').setFontColor('#ffffff');
  [ '1) Escribí una fila nueva abajo de todo: Activa = SI, un Nombre, la URL y el Parser.',
    '2) Parser: probá "wpjson" primero (sirve en cualquier sitio hecho con WordPress).',
    '   Si no trae nada, cambialo a "generico". Los demás son a medida de cada sitio.',
    '3) Mañana el sistema la busca sola y te completa las columnas azules.',
    '',
    'COLORES: verde = anduvo bien · rojo = dio error · naranja = hace 3 corridas no trae nada',
    'gris = apagada · celeste = agregada por ustedes, todavía sin probar',
    '',
    'Si una fuente queda en rojo o naranja, entrá a esa web a mano hasta que se arregle.'
  ].forEach(function (t, i) {
    sh.getRange(f + 1 + i, 1, 1, 8).merge().setValue(t).setFontColor(i >= 5 ? '#666666' : '#000000');
  });

  var anchos = [60, 240, 330, 95, 260, 120, 95, 95, 110, 300, 120];
  for (var c = 0; c < anchos.length; c++) sh.setColumnWidth(c + 1, anchos[c]);
  sh.setFrozenRows(1);
  return { total: filas.length, conError: conError };
}

function norUrl_(u) {
  return (u || '').toString().toLowerCase().replace(/^https?:\/\//, '').replace(/\/+$/, '').trim();
}

function bajarDiagnostico_() {
  // OJO: acá faltaba el respaldo y por eso la hoja Fuentes quedaba sin estados.
  var t = pedir_(CFG.DIAG + '?t=' + Date.now(), {});
  if (!t) t = pedir_(CFG.DIAG_RESPALDO, { 'Accept': 'application/vnd.github.raw' });
  if (!t) return [];
  try {
    var tabla = Utilities.parseCsv(t);
    if (tabla.length < 2) return [];
    var cab = tabla[0], out = [];
    for (var i = 1; i < tabla.length; i++) {
      var o = {};
      for (var c = 0; c < cab.length; c++) o[cab[c]] = tabla[i][c];
      out.push(o);
    }
    return out;
  } catch (e) { return []; }
}

// ====================== ORDENAR LA PLANILLA ===========================
/**
 * Borra las hojas viejas que quedaron de versiones anteriores.
 * Pregunta antes de borrar nada.
 */
function limpiarHojasViejas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quedan = [CFG.HOJA, CFG.RESUMEN, CFG.FUENTES, 'Mails'];
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

var VERSION = 'v5 · 09/09/2026';

/**
 * Revisa si la planilla puede leer los datos y muestra el resultado.
 * Sirve para saber si algo falla y qué exactamente.
 */
function probarConexion() {
  var lineas = ['VERSIÓN DEL SCRIPT: ' + VERSION, ''];
  function probar(nombre, url, headers) {
    try {
      var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: headers || {} });
      var code = r.getResponseCode();
      var txt = code === 200 ? r.getContentText() : '';
      var filas = txt ? Utilities.parseCsv(txt).length - 1 : 0;
      lineas.push((code === 200 ? '✓ ' : '✗ ') + nombre + ': HTTP ' + code +
                  (code === 200 ? '  (' + filas + ' filas)' : ''));
      return filas;
    } catch (e) {
      lineas.push('✗ ' + nombre + ': ' + e);
      return 0;
    }
  }
  var lic  = probar('Listado de licitaciones', CFG.CSV + '?t=' + Date.now());
  if (!lic) probar('  ...respaldo por la API', CFG.CSV_RESPALDO, { 'Accept': 'application/vnd.github.raw' });
  var diag = probar('Estado de las fuentes',  CFG.DIAG + '?t=' + Date.now());
  if (!diag) probar('  ...respaldo por la API', CFG.DIAG_RESPALDO, { 'Accept': 'application/vnd.github.raw' });

  lineas.push('');
  if (lic && diag) lineas.push('TODO BIEN. Usá "Actualizar ahora" y las fuentes se van a pintar.');
  else if (lic && !diag) lineas.push('Se leen las licitaciones pero NO el estado de las fuentes:\npor eso quedan en gris. Avisale a Claude.');
  else lineas.push('No se puede leer nada. Puede ser un corte momentáneo:\nprobá de nuevo en unos minutos.');

  SpreadsheetApp.getUi().alert('Prueba de conexión', lineas.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function irAFuentes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.FUENTES);
  if (!sh) { actualizar(); sh = ss.getSheetByName(CFG.FUENTES); }
  if (sh) { sh.activate(); ss.toast('Verde = anduvo · Rojo = error · Naranja = hace días no trae nada', 'Fuentes', 10); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏗️ Licitaciones')
    .addItem('Actualizar ahora', 'actualizar')
    .addSeparator()
    .addItem('Revisar estado de las fuentes', 'irAFuentes')
    .addItem('Probar conexión (si algo no anda)', 'probarConexion')
    .addSeparator()
    .addItem('Borrar hojas viejas', 'limpiarHojasViejas')
    .addItem('Instalar / reinstalar automático', 'instalarTodo')
    .addToUi();
}
