/**
 * AccountStore · App del celular para "Planilla de Suscripciones 2.0"
 *
 * Web app de Google Apps Script (proyecto independiente). Lee y escribe
 * directo en la planilla, sin cambiar su estructura:
 *  - BasedeDatos: solo escribe en columnas manuales (Nombre, ID, Perfil,
 *    Fecha, Monto, Teléfono, Correo, Servicios, Dispositivos, PIN, PINS,
 *    MESES). Vencimiento, Días, ESTADO y SMS siguen siendo fórmulas.
 *  - Hojas de servicio: solo escribe el nombre en la columna "Usuario" de
 *    un lugar, y la contraseña / último cambio de una cuenta.
 */

const CONFIG = {
  // PLANILLA PRINCIPAL. Para probar algo en la copia de prueba, reemplazá por:
  // 1it5CN1YWFWt_Pb9rsawHr7bmdoTebtTRlJaT1vK2vzE
  SPREADSHEET_ID: '1tTuAEALmDEjCld0UtxtdCuuC1mwPuOLtilAmPsYW9AU',
  HOJA_CLIENTES: 'BasedeDatos',
  HOJA_REGISTRO: 'RegistrodeDatos',
  RANGO_LEYENDA: 'L2:O17',          // servicio / emoji en RegistrodeDatos
  // Hojas que nunca se tratan como servicios (las ocultas ya se ignoran)
  EXCLUIR: ['RegistrodeDatos', 'BasedeDatos', 'TRANSFERENCIAS', 'CODE',
            'PAGOPAR', 'INFO', 'Facturas', 'Inicio', 'Updates-Ventas', 'RegistroApp'],
  CODIGO_PAIS: '595',

  // Ocupación automática de las cuentas: "(LLENO)" o "(+N)" según los perfiles libres.
  // Máximo de perfiles por cuenta; los servicios que no están acá usan la cantidad de filas de usuario.
  CAPACIDAD: { 'Netflix': 5, 'HBO MAX': 5, 'Crunchyroll': 5, 'Disney': 6, 'Prime Video': 5, 'Paramount': 4, 'CapCut': 3, 'FLUJO TV': 3 },
  SIN_OCUPACION: ['Youtube Premium', 'FLOW TV', 'Spotify', 'AppleTV'],   // hojas donde no se calcula ni se escribe la ocupación (AppleTV: sin límite)
  VIP_NO_OCUPAN_EN: ['Disney', 'ChatGPT'],   // los clientes VIP ocupan lugar como todos, salvo en estas hojas
  PRIVADA_OCUPA_TODO: ['Crunchyroll'],     // "Crunchyroll-Privada": ese cliente ocupa todos los perfiles de la cuenta
  MAX_FILAS_OCUPACION: 12,   // bloques más grandes son listas, no cuentas: no llevan marca
  // Mientras esté en false, no escribe nada: ejecutá revisarOcupacion() para ver qué cambiaría.
  OCUPACION_AUTOMATICA: true,
  // Horario del proceso automático de cada 30 minutos (Pendientes y ocupación): desde / hasta (24 = medianoche)
  HORARIO_AUTOMATICO: [9, 24],

  // Transferencias: tu web app que lee los correos del banco (la misma que dispara el tilde de H1)
  TRANSFERENCIAS_URL: 'https://script.google.com/macros/s/AKfycbzaonTqTdI6utDZLx9SK1jXPfJbvghxR1qAyAjXCk8YdUWmeukx8m3cl8l3m9Ni5WIGCg/exec',
  HOJA_TRANSFERENCIAS: 'TRANSFERENCIAS',

  // Quién puede entrar a la app. Los PIN NO van acá: se guardan en
  // Configuración del proyecto → Propiedades de la secuencia de comandos, una por persona:
  //   PIN_MARCELO_BENITEZ = 6 números, PIN_ADRIAN_BENITEZ = ..., PIN_NATALIA_BENITEZ = ...
  // (Si escribís un PIN acá y ejecutás probar, se guarda solo en la configuración.)
  USUARIOS: {
    'Marcelo Benitez': '',   // admin
    'Adrian Benitez': '',    // supervisor
    'Natalia Benitez': ''    // pagos, fechas y vencimientos
  },
  // Cómo se muestra el rol de cada uno en el menú de la app
  ROLES: {
    'Marcelo Benitez': 'Administrador',
    'Adrian Benitez': 'Supervisor',
    'Natalia Benitez': 'Pagos y vencimientos'
  },
  // Lo que NO puede hacer cada uno. Quien no aparece acá puede hacer todo.
  // Opciones: 'cambiarPass' (contraseñas de cuentas), 'liberarLugar', 'darDeBaja'
  RESTRICCIONES: {
    'Natalia Benitez': ['cambiarPass', 'liberarLugar', 'darDeBaja', 'destrabarPago']
  },
  DIAS_SESION: 30,                  // cada cuántos días vuelve a pedir el PIN
  // A quién avisar por correo si alguien erra el PIN varias veces seguidas.
  // Vacío = al dueño del proyecto (tu Gmail).
  EMAIL_ALERTAS: '',

  // Servicios que no llevan dispositivo: si el cliente solo tiene estos, Dispositivos queda en "-"
  SIN_DISPOSITIVO: ['Spotify', 'ChatGPT', 'Office', 'Flujo', 'Youtube', 'Canva', 'CapCut', 'Deezer',
                    'Tidal', 'Apple Music', 'Google One', 'Game Pass', 'Xbox Game Pass', 'Twitch', 'Netflix-Privado'],
  // Servicios que sí o sí llevan el correo del cliente
  CON_CORREO: ['Netflix-Privado', 'Office', 'Spotify', 'Youtube', 'Canva', 'Deezer', 'Tidal', 'Apple Music',
               'Duolingo', 'Google One', 'Twitch', 'Game Pass', 'Xbox Game Pass'],                  // cada cuántos días vuelve a pedir el PIN
  HOJA_LOG: 'RegistroApp',          // hoja donde se anota cada cambio (se crea sola)
  // La columna Perfil de BasedeDatos es una fórmula que suma los perfiles de
  // la columna Servicios. La app no la escribe, solo la muestra.
  // Poné false si algún día volvés a cargarla a mano y querés que la app la complete.
  PERFIL_ES_FORMULA: true,
  VACIO: '-',                       // valor para campos opcionales vacíos
  // Al renovar, la app calcula el nuevo vencimiento (EDATE) en vez de esperar a que la planilla recalcule.
  // Comprobalo una vez con revisarCalculoVencimiento(); si diera diferencias, poné false.
  RESPUESTA_RAPIDA: true,
  // Bajas: columnas "Baja" (fecha en que se canceló) y "Venció" (vencimiento que tenía), después de MESES.
  // Se crean con prepararColumnasBaja() y se llenan solas (desde la app, o con anotarBajasManuales() al poner X a mano).
  DIAS_CONTINUAR: 45,                 // si vuelve dentro de estos días, la app sugiere seguir desde el vencimiento anterior
  MORA_BAJA: { 'Tigo Sports': 5000 }, // recargo si se lo da de baja después de esperarle días: en Monto queda -5000

  // Mensajes de WhatsApp. Variables: {nombre} {servicios} {vencimiento} {monto}
  MSG_POR_VENCER: 'Hola {nombre}! 👋\nTe recuerdo que tu suscripción de *{servicios}* vence el *{vencimiento}*.\n_*📌Recuerda abonar para evitar cortes👏*_',
  MSG_VENCIDO: 'Hola {nombre}! 👋\nTu suscripción de *{servicios}* venció el *{vencimiento}*.\n_*📌Recuerda abonar para evitar cortes👏*_',
  MSG_RENOVADO: '¡Listo {nombre}! ✅\nTu suscripción de *{servicios}* quedó renovada hasta el *{vencimiento}*.\n¡Gracias por elegirnos! 💯'
};

// Encabezados de BasedeDatos (normalizados) -> campo
const CAMPOS_CLIENTE = {
  nombre: ['nombre y apellido'],
  id: ['id'],
  perfil: ['perfil', 'perfiles', 'perfil es'],
  fecha: ['fecha'],
  monto: ['monto'],
  telefono: ['telefono'],
  correo: ['correo'],
  servicios: ['servicios', 'servicio s'],
  dispositivos: ['dispositivos', 'dispositivo s'],
  pin: ['pin'],
  pins: ['pins'],
  vencimiento: ['vencimiento'],
  dias: ['dias'],
  estado: ['estado'],
  meses: ['meses'],
  baja: ['baja', 'fecha baja', 'fecha de baja'],
  vencio: ['vencio', 'vencia', 'venc anterior', 'vencimiento anterior']
};
const EDITABLES = ['nombre', 'id', 'perfil', 'fecha', 'monto', 'telefono', 'correo',
                   'servicios', 'dispositivos', 'pin', 'pins', 'meses', 'baja', 'vencio'];   // baja y vencio: solo los escribe el servidor
// Campos que la app puede escribir en esta planilla
function editables_() {
  return CONFIG.PERFIL_ES_FORMULA ? EDITABLES.filter(function (k) { return k !== 'perfil'; }) : EDITABLES;
}

/* ------------------------------------------------------------------ */
/* Web app y acceso                                                    */
/* ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AccountStore')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * API para la app instalada (GitHub Pages): recibe {fn, args} y devuelve {ok, r} o {ok, error}.
 * Solo se pueden llamar estas funciones, y todas menos "entrar" exigen la sesión del PIN.
 */
function doPost(e) {
  const API = {
    entrar: entrar, getInicial: getInicial, getBajas: getBajas, getServiciosLista: getServiciosLista,
    getServicio: getServicio, ubicarCliente: ubicarCliente, renovarCliente: renovarCliente,
    guardarCliente: guardarCliente, darDeBaja: darDeBaja, nuevoCliente: nuevoCliente,
    asignarLugar: asignarLugar, liberarLugar: liberarLugar, cambiarPassCuenta: cambiarPassCuenta,
    getPendientes: getPendientes, liberarVarios: liberarVarios, getResumen: getResumen, getCuenta: getCuenta, getCliente: getCliente,
    getTransferencias: getTransferencias, cargarTransferencias: cargarTransferencias, actualizarOcupacion: actualizarOcupacion,
    vincularPago: vincularPago, destrabarPago: destrabarPago, quitarVinculo: quitarVinculo, buscarPedidoPagopar: buscarPedidoPagopar, getOtrosPagos: getOtrosPagos
  };
  let salida;
  try {
    const pedido = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = Object.prototype.hasOwnProperty.call(API, pedido.fn) ? API[pedido.fn] : null;
    if (!fn) throw new Error('Pedido no válido.');
    salida = { ok: true, r: fn.apply(null, Array.isArray(pedido.args) ? pedido.args : []) };
  } catch (err) {
    salida = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

/** Ejecutala desde el editor para autorizar y probar la conexión. */
function probar() {
  pasarPinesAConfiguracion_();
  problemasUsuarios_().forEach(function (p) { Logger.log('ATENCIÓN: ' + p); });
  const usuarios = usuariosValidos_();
  if (!usuarios.length) {
    Logger.log('ATENCIÓN: nadie tiene PIN todavía. Agregalos en Configuración del proyecto (PIN_NOMBRE_APELLIDO) o escribilos en USUARIOS y ejecutá probar.');
    return;
  }
  const t = firmar_(usuarios[0]);
  const d = getInicial(t);
  const bajas = getBajas(t);
  const hojas = getServiciosLista(t, true);
  Logger.log('Usuarios con acceso: %s', usuarios.join(', '));
  usuarios.forEach(function (u) {
    const no = Object.keys(PERMISOS_).filter(function (k) { return !permisos_(u)[k]; }).map(function (k) { return PERMISOS_[k]; });
    Logger.log('  %s: %s', u, no.length ? 'no puede ' + no.join(', ') : 'puede todo');
  });
  Logger.log('Activos: %s | Dados de baja: %s | Hojas de servicio: %s', d.clientes.length, bajas.length, hojas.join(', '));
}

/** Si hay PIN escritos en USUARIOS, los guarda en la configuración del proyecto (una sola vez). */
function pasarPinesAConfiguracion_() {
  const props = PropertiesService.getScriptProperties();
  const guardadas = props.getProperties(), nuevas = {};
  Object.keys(CONFIG.USUARIOS || {}).forEach(function (u) {
    const enCodigo = String(CONFIG.USUARIOS[u] || '').trim(), clave = clavePin_(u);
    if (!enCodigo) {
      Logger.log('%s: PIN %s', u, guardadas[clave] ? 'guardado en la configuración del proyecto ✓' : 'NO cargado');
      return;
    }
    if (!/^\d{6}$/.test(enCodigo)) { Logger.log('%s: el PIN escrito en el código no tiene 6 números; no se guardó.', u); return; }
    if (guardadas[clave] && guardadas[clave] !== enCodigo) {
      Logger.log('%s: el PIN del código es distinto al guardado. Vale el guardado (%s); si querés cambiarlo, editalo en Configuración del proyecto.', u, clave);
      return;
    }
    if (!guardadas[clave]) nuevas[clave] = enCodigo;
    Logger.log('%s: PIN guardado en la configuración del proyecto ✓. Ya podés borrarlo de Code.gs.', u);
  });
  if (Object.keys(nuevas).length) props.setProperties(nuevas);
  olvidarAuth_();   // así un PIN nuevo o cambiado vale en el momento
}

/** Pantalla de PIN: devuelve una sesión firmada si el PIN es correcto. */
const MAX_FALLOS_ = 8;       // intentos errados seguidos antes de trabar la entrada 15 minutos

function entrar(pin) {
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('pinFallos') || 0);
  if (fallos >= MAX_FALLOS_) throw new Error('Demasiados intentos. Esperá 15 minutos.');
  const p = String(pin || '').trim();
  const usuario = usuariosValidos_().filter(function (u) { return pinDe_(u) === p; })[0];
  if (!usuario) {
    const n = fallos + 1;
    cache.put('pinFallos', String(n), 900);
    registrar_(null, '(desconocido)', 'PIN incorrecto', '', 'intento ' + n + ' de ' + MAX_FALLOS_);
    if (n >= MAX_FALLOS_) avisarBloqueo_();
    Utilities.sleep(700);
    throw new Error(usuariosValidos_().length ? 'PIN incorrecto.' : 'Todavía no hay PIN configurados en la app.');
  }
  cache.remove('pinFallos');
  return { token: firmar_(usuario), usuario: usuario };
}

/* ------------------------------------------------------------------ */
/* Lecturas (todas piden la sesión como primer dato)                   */
/* ------------------------------------------------------------------ */

/**
 * Carga rápida: solo clientes que no están dados de baja (unos 1.000 de 4.400).
 * Los dados de baja llegan después con getBajas() y la lista de hojas con
 * getServiciosLista(), para que la app se pueda usar mientras tanto.
 */
function getInicial(token) {
  const quien = usuario_(token);
  const t0 = Date.now();
  const { ss, tz } = ctx_();
  const h = hojaClientes_(ss);
  const t1 = Date.now();
  const r = leerClientes_(h, tz, false);
  registrarIngreso_(ss, quien, tz);
  console.log('getInicial (%s): abrir %s ms, leer %s ms, armar %s ms, %s activos',
    quien, t1 - t0, r.tLeer, r.tArmar, r.lista.length);
  return {
    usuario: quien,
    rol: (CONFIG.ROLES || {})[quien] || '',
    permisos: permisos_(quien),
    reglas: { sinDispositivo: CONFIG.SIN_DISPOSITIVO, conCorreo: CONFIG.CON_CORREO, mora: CONFIG.MORA_BAJA || {}, diasContinuar: CONFIG.DIAS_CONTINUAR || 45 },
    hoy: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    clientes: r.lista,
    leyenda: leyenda_(ss),
    perfilAuto: !!CONFIG.PERFIL_ES_FORMULA,
    msgs: { porVencer: CONFIG.MSG_POR_VENCER, vencido: CONFIG.MSG_VENCIDO, renovado: CONFIG.MSG_RENOVADO }
  };
}

/** Los clientes dados de baja (Fecha = X). Se piden en segundo plano. */
function getBajas(token) {
  usuario_(token);
  const t0 = Date.now();
  const { ss, tz } = ctx_();
  const h = hojaClientes_(ss);
  const r = leerClientes_(h, tz, true);
  console.log('getBajas: total %s ms, %s bajas', Date.now() - t0, r.lista.length);
  return r.lista;
}

/** Nombres de las hojas de servicio (guardados 6 horas). */
function getServiciosLista(token, forzar) {
  usuario_(token);
  const t0 = Date.now();
  const { ss } = ctx_();
  const lista = hojasServicio_(ss, !!forzar);
  console.log('getServiciosLista: %s ms, %s hojas%s', Date.now() - t0, lista.length, forzar ? ' (recalculada)' : '');
  return lista;
}

function leerClientes_(h, tz, bajas) {
  const t0 = Date.now();
  const last = ultimaFila_(h.sh, h.cols.nombre);
  // Velocidad: solo hasta la última columna que usa la app (las de más a la derecha, como SMS, son textos largos que no se usan)
  let ancho = 1;
  Object.keys(h.cols).forEach(function (k) { ancho = Math.max(ancho, h.cols[k] + 1); });
  ancho = Math.min(ancho, h.ncol);
  const vals = last > 1 ? h.sh.getRange(2, 1, last - 1, ancho).getValues() : [];
  const t1 = Date.now();
  const lista = [];
  vals.forEach(function (v, i) {
    if (String(v[h.cols.nombre]).trim() === '') return;
    const esBaja = String(v[h.cols.fecha]).trim().toUpperCase() === 'X';
    if (esBaja === bajas) lista.push(enc_(v, i + 2, h.cols, tz));
  });
  return { lista: lista, tLeer: t1 - t0, tArmar: Date.now() - t1 };
}

function getServicio(token, hoja) {
  usuario_(token);
  const t0 = Date.now();
  const { ss } = ctx_();
  const r = servicio_(ss, hoja);
  try { calcularOcupacion_(hoja, r, vipsGuardados_()); } catch (e) { }   // disponibles de cada cuenta, para la app
  console.log('getServicio %s: %s ms, %s cuentas', hoja, Date.now() - t0, r.cuentas.length);
  return r;
}

/** Dónde aparece un cliente en las hojas de servicio. */
/**
 * "Ver en qué cuentas está": busca al cliente en la columna A de las hojas.
 *  - Por defecto, solo en las hojas de sus servicios (Netflix-Premium → Netflix; HBO MAX-2 → HBO MAX).
 *  - Si no aparece ahí, o es cliente VIP, o se pide "todas", busca en todas las hojas.
 * Con Google Sheets API lee todas las columnas A en un pedido y las filas de datos en otro.
 */
function ubicarCliente(token, nombre, servicios, todas) {
  usuario_(token);
  const t0 = Date.now();
  const { ss } = ctx_();
  const buscado = String(nombre).trim();
  const lista = hojasServicio_(ss, false);
  const suyas = todas ? null : hojasDeServicios_(servicios, lista);
  let enTodas = !suyas || !suyas.length;
  let res = buscarEnHojas_(ss, enTodas ? lista : suyas, buscado);
  if (!res.length && !enTodas) { enTodas = true; res = buscarEnHojas_(ss, lista, buscado); }
  console.log('ubicarCliente: %s hojas revisadas, %s lugares, %s ms', enTodas ? lista.length : suyas.length, res.length, Date.now() - t0);
  return { lugares: res, todas: enTodas, revisadas: enTodas ? lista.length : suyas.length };
}

// Servicio → hoja. Equivalencias que no salen del nombre:
const ALIAS_SERVICIO_ = { youtubemusic: 'youtubepremium' };
function claveHoja_(s) { return norm_(s).replace(/[^a-z0-9]/g, ''); }
/** Hojas de los servicios del cliente; null si hay que buscar en todas (clientes VIP). */
function hojasDeServicios_(servicios, lista) {
  const out = [];
  const tokens = String(servicios || '').split(',').map(function (t) { return t.trim(); }).filter(function (t) { return t && t !== '-'; });
  for (let i = 0; i < tokens.length; i++) {
    let k = claveHoja_(tokens[i].replace(/-\d+\s*$/, '').split('-')[0]);   // "Netflix-Premium-2" → netflix
    if (!k) continue;
    if (ALIAS_SERVICIO_[k]) k = ALIAS_SERVICIO_[k];
    if (k === 'vip') return null;
    const h = lista.filter(function (x) { return claveHoja_(x) === k; })[0] ||
              lista.filter(function (x) { const c = claveHoja_(x); return c.indexOf(k) === 0 || k.indexOf(c) === 0; })[0];
    if (h && out.indexOf(h) < 0) out.push(h);
  }
  return out;
}

function buscarEnHojas_(ss, hojas, buscado) {
  const colas = leerColumnasA_(ss, hojas);
  const res = [], faltan = [];
  const esCuenta = function (s) { return /^cuenta\b/.test(norm_(s || '')); };
  hojas.forEach(function (hoja) {
    const colA = colas[hoja] || [];
    colA.forEach(function (v, i) {
      if (v !== buscado) return;
      let h = i - 1;
      while (h >= 0 && norm_(colA[h]) !== 'usuario') h--;
      if (h < 0) return;
      const cuenta = esCuenta(colA[h - 1]) ? colA[h - 1] : (esCuenta(colA[h - 2]) ? colA[h - 2] : '');
      const plan = esCuenta(colA[h - 1]) ? '' : (colA[h - 1] || '');
      const lugar = { hoja: hoja, row: i + 1, header: h + 1, cuenta: cuenta, plan: plan, inactiva: norm_(plan) === 'miembro extra' };
      res.push(lugar);
      if (!lugar.inactiva && h >= 1 && !esCuenta(colA[h - 1])) faltan.push({ lugar: lugar, fila: h });   // fila de datos de la cuenta
    });
  });
  // ¿alguna de esas cuentas dice "reactivar"? (todas las filas de datos en un solo pedido)
  if (faltan.length) {
    const filas = leerFilas_(ss, faltan.map(function (f) { return { hoja: f.lugar.hoja, row: f.fila }; }));
    faltan.forEach(function (f, k) {
      f.lugar.inactiva = (filas[k] || []).some(function (x) { return /reactivar/i.test(String(x)); });
    });
  }
  return res;
}
const rangoHoja_ = function (h) { return "'" + String(h).replace(/'/g, "''") + "'"; };
const hayApiSheets_ = function () { return typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values; };
/** Columna A (como texto) de varias hojas: { hoja: [valores] } */
function leerColumnasA_(ss, hojas) {
  const out = {};
  if (hayApiSheets_()) {
    try {
      const r = Sheets.Spreadsheets.Values.batchGet(CONFIG.SPREADSHEET_ID, { ranges: hojas.map(function (h) { return rangoHoja_(h) + '!A:A'; }), majorDimension: 'ROWS' });
      (r.valueRanges || []).forEach(function (vr, i) { out[hojas[i]] = (vr.values || []).map(function (f) { return String(f[0] == null ? '' : f[0]).trim(); }); });
      return out;
    } catch (e) { console.warn('Lectura en un solo pedido falló, sigo hoja por hoja: ' + e); }
  }
  hojas.forEach(function (h) {
    const sh = ss.getSheetByName(h);
    const n = sh ? sh.getLastRow() : 0;
    out[h] = n < 1 ? [] : sh.getRange(1, 1, n, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); });
  });
  return out;
}
/** Filas sueltas (columnas A a L, como texto): [[...], ...] en el mismo orden que el pedido */
function leerFilas_(ss, pedidos) {
  if (hayApiSheets_()) {
    try {
      const r = Sheets.Spreadsheets.Values.batchGet(CONFIG.SPREADSHEET_ID, { ranges: pedidos.map(function (p) { return rangoHoja_(p.hoja) + '!A' + p.row + ':L' + p.row; }), majorDimension: 'ROWS' });
      return (r.valueRanges || []).map(function (vr) { return (vr.values && vr.values[0]) || []; });
    } catch (e) { console.warn('Lectura en un solo pedido falló, sigo fila por fila: ' + e); }
  }
  return pedidos.map(function (p) {
    const sh = ss.getSheetByName(p.hoja);
    return sh ? sh.getRange(p.row, 1, 1, Math.min(COLS_RECORRIDO, sh.getLastColumn())).getDisplayValues()[0] : [];
  });
}

/* ------------------------------------------------------------------ */
/* Escrituras en BasedeDatos (cada una queda anotada en RegistroApp)   */
/* ------------------------------------------------------------------ */

/**
 * Renueva respetando cómo se usa la planilla: Vencimiento = EDATE(Fecha, MESES).
 *  modo 'sumar': deja la Fecha de activación y suma meses a MESES (sigue desde el vencimiento).
 *  modo 'desde': empieza de nuevo: Fecha = fechaISO y MESES = meses (reactivar o reiniciar).
 */
function renovarCliente(token, row, nombre, modo, meses, fechaISO, monto, esperado, pago) {
  const t0 = Date.now();
  const T = {};
  const quien = usuario_(token);
  T.sesion = Date.now() - t0;
  try {
  const tc = Date.now();
  return conLock_(function () {
    T.candado = Date.now() - tc;
    const tl = Date.now();
    const { ss, tz } = ctx_();
    const h = hojaClientes_(ss);
    const n = Math.round(Number(meses));
    if (!(n >= 1 && n <= 60)) throw new Error('La cantidad de meses tiene que estar entre 1 y 60.');
    const fr = filaYValores_(h, row, nombre), r = fr.r, fila = fr.fila;
    const antes = enc_(fila, r, h.cols, tz);
    verificarSinCambios_(antes, esperado);
    verificarPago_(pago, antes[1]);   // el pago (transferencia o número de referencia) no puede estar usado para otro cliente
    let datos;
    if (modo === 'sumar') {
      const fecha = fila[h.cols.fecha], actuales = Number(fila[h.cols.meses]);
      if (!(fecha instanceof Date) || isNaN(actuales)) {
        throw new Error('Este cliente no tiene fecha de activación. Usá "Empezar de nuevo".');
      }
      datos = { meses: actuales + n };
    } else {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fechaISO))) throw new Error('Elegí la fecha de inicio.');
      datos = { fecha: fechaISO, meses: n };
    }
    if (monto !== '' && monto != null) datos.monto = monto;
    if (antes[4].toUpperCase() === 'X' && h.cols.baja != null) { datos.baja = ''; datos.vencio = ''; }   // reactivado: se limpian Baja y Venció
    T.leer = Date.now() - tl;
    const tw = Date.now();
    const cambios = escribirCampos_(h, r, datos, tz, fila);
    // Velocidad: el vencimiento y los días se calculan acá (Vencimiento = EDATE(Fecha; MESES)) en vez de esperar
    // a que la planilla recalcule todas sus fórmulas para leerlos (eso era lo más lento y lo que variaba).
    let despues = CONFIG.RESPUESTA_RAPIDA ? despuesCalculado_(antes, datos, tz) : null;
    if (!despues) { SpreadsheetApp.flush(); T.escribir = Date.now() - tw; const tr = Date.now(); despues = leerCliente_(h, r, tz); T.releer = Date.now() - tr; }
    else {
      // Comprobación: se vuelve a leer SOLO Fecha y MESES (celdas sin fórmula) para confirmar que quedaron escritas.
      // Si no coinciden, no se da por guardado: así la app nunca dice "renovado" si la planilla no cambió.
      SpreadsheetApp.flush();
      T.escribir = Date.now() - tw;
      const tv = Date.now();
      const fReal = h.sh.getRange(r, h.cols.fecha + 1).getValue(), mReal = h.sh.getRange(r, h.cols.meses + 1).getValue();
      const fTxt = fReal instanceof Date ? Utilities.formatDate(fReal, tz, 'yyyy-MM-dd') : String(fReal).trim();
      T.releer = Date.now() - tv;
      if (fTxt !== String(despues[4]) || Number(mReal) !== Number(despues[15])) {
        console.warn('renovarCliente: la planilla no quedó como se esperaba (fila %s): fecha %s / meses %s, esperado %s / %s', r, fTxt, mReal, despues[4], despues[15]);
        throw new Error('La planilla no guardó el cambio (fila ' + r + ': quedó fecha ' + (fTxt || 'vacía') + ', meses ' + mReal +
          '). No se dio por renovado. Tocá Actualizar y avisale al administrador.');
      }
    }
    const accion = antes[4].toUpperCase() === 'X' ? 'Reactivó' : modo === 'sumar' ? 'Renovó +' + n + (n === 1 ? ' mes' : ' meses') : 'Reinició desde ' + isoDmy_(fechaISO);
    const tp = Date.now();
    if (pago) marcarPagoUsado_(pago, despues[1], quien);
    T.pago = Date.now() - tp;
    const tg = Date.now();
    registrar_(ss, quien, accion, despues[1], detalleVenc_(antes, despues, cambios) + (pago ? ' · pagó con ' + textoPago_(pago) : ''));
    T.registro = Date.now() - tg;
    return despues;
  });
  } finally {
    // aparece en Apps Script → Ejecuciones: cuánto tardó cada paso
    console.log('renovarCliente: %s ms en total | esperar turno %s, sesión %s, leer %s, escribir %s, releer fórmulas %s, pago %s, registro %s',
      Date.now() - t0, T.candado, T.sesion, T.leer, T.escribir, T.releer, T.pago, T.registro);
  }
}

/**
 * Cómo queda el cliente después de renovar, sin volver a leer la planilla:
 * Vencimiento = EDATE(Fecha; MESES) y Días = Vencimiento − hoy. Devuelve null si no se puede calcular
 * (entonces se lee de la planilla como antes). Para comprobar que coincide con tus fórmulas: revisarCalculoVencimiento().
 */
function despuesCalculado_(antes, datos, tz) {
  try {
    const d = antes.slice();
    if (datos.fecha != null) d[4] = String(datos.fecha);
    if (datos.meses != null) d[15] = Number(datos.meses);
    if (datos.monto != null) d[5] = valorCampo_('monto', datos.monto, tz);
    if (datos.baja != null) d[16] = String(datos.baja);
    if (datos.vencio != null) d[17] = String(datos.vencio);
    const venc = edateIso_(d[4], Number(d[15]));
    if (!venc) return null;
    d[12] = venc;
    d[13] = difDias_(venc, Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));
    d[14] = estadoPorDias_(d[13]);
    return d;
  } catch (e) { return null; }
}
/** Estado como la fórmula de la columna ESTADO: más de 3 días ACTIVO, menos de 1 VENCIDO, si no A VENCER. */
function estadoPorDias_(d) { return d > 3 ? 'ACTIVO' : d < 1 ? 'VENCIDO' : 'A VENCER'; }
/** Perfil como la fórmula de BasedeDatos: suma el "-N" del final de cada servicio (sin número = 1). */
function perfilDeServicios_(nombre, servicios) {
  const sv = String(servicios || '').trim();
  if (!String(nombre || '').trim() || !sv || sv === '-') return '';
  let n = 0;
  sv.split(',').map(function (x) { return x.trim(); }).filter(String).forEach(function (x) {
    const m = x.match(/-(\d+)$/); n += m ? Number(m[1]) : 1;
  });
  return String(n);
}
/** Cómo queda el cliente después de editar (sin releer la planilla). null si no se puede calcular. */
function despuesDeCambios_(antes, cambios, tz) {
  try {
    const IDX = { nombre: 1, id: 2, perfil: 3, fecha: 4, monto: 5, telefono: 6, correo: 7, servicios: 8, dispositivos: 9, pin: 10, pins: 11, meses: 15, baja: 16, vencio: 17 };
    const d = antes.slice();
    let fechaOMeses = false, servicios = false;
    cambios.forEach(function (c) {
      const i = IDX[c.campo];
      if (i == null) return;
      let v = c.valor;
      if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      else if (c.campo === 'telefono') v = v === '' ? '' : String(Math.round(Number(v)));
      else if (typeof v !== 'number') v = String(v == null ? '' : v).trim();
      d[i] = v;
      if (c.campo === 'fecha' || c.campo === 'meses') fechaOMeses = true;
      if (c.campo === 'servicios' || c.campo === 'nombre') servicios = true;
    });
    if (servicios && CONFIG.PERFIL_ES_FORMULA) d[3] = perfilDeServicios_(d[1], d[8]);
    if (fechaOMeses) {
      if (String(d[4]).toUpperCase() === 'X') { d[12] = 'CANCELADO'; d[13] = null; d[14] = 'INACTIVO'; }
      else {
        const v = edateIso_(d[4], Number(d[15]));
        if (!v) return null;
        d[12] = v; d[13] = difDias_(v, Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')); d[14] = estadoPorDias_(d[13]);
      }
    }
    return d;
  } catch (e) { return null; }
}
/** Relee SOLO las celdas que se cambiaron (no tienen fórmula, así que es rápido). Si alguna no quedó, corta con error. */
function confirmarEscrito_(h, r, cambios, tz) {
  if (!cambios.length) return;
  SpreadsheetApp.flush();
  const malos = cambios.filter(function (c) {
    const real = h.sh.getRange(r, h.cols[c.campo] + 1).getValue();
    return !iguales_(real, c.valor, tz);
  });
  if (malos.length) {
    console.warn('No quedaron escritos en la fila %s: %s', r, malos.map(function (c) { return c.campo; }).join(', '));
    throw new Error('La planilla no guardó ' + malos.map(function (c) { return ETIQUETAS_[c.campo] || c.campo; }).join(', ') +
      ' (fila ' + r + '). Tocá Actualizar y revisá antes de volver a guardar.');
  }
}
/** EDATE de Sheets: misma fecha N meses después; si ese mes no tiene el día, el último día del mes. */
function edateIso_(iso, meses) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m || !(meses >= 0)) return '';
  const total = Number(m[2]) - 1 + Math.round(meses);
  const y = Number(m[1]) + Math.floor(total / 12), mo = total % 12;
  const ultimo = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
  const d = Math.min(Number(m[3]), ultimo);
  return y + '-' + String(mo + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}
function difDias_(a, b) {
  const p = function (s) { const x = s.split('-').map(Number); return Date.UTC(x[0], x[1] - 1, x[2]); };
  return Math.round((p(a) - p(b)) / 864e5);
}
/**
 * Ejecutala UNA VEZ desde el editor: compara, para todos los clientes activos, el vencimiento y los días
 * que calcula la app con los que muestran tus fórmulas. Si dice "0 diferencias", la respuesta rápida es exacta.
 * Si hay diferencias, poné RESPUESTA_RAPIDA: false en CONFIG y avisame.
 */
function revisarCalculoVencimiento() {
  const { ss, tz } = ctx_();
  const h = hojaClientes_(ss);
  const lista = leerClientes_(h, tz, false).lista;
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  let revisados = 0, malos = 0;
  lista.forEach(function (c) {
    if (!/^\d{4}-/.test(c[4]) || typeof c[15] !== 'number' || !/^\d{4}-/.test(c[12])) return;
    revisados++;
    const v = edateIso_(c[4], c[15]), dd = difDias_(v, hoy);
    if (v !== c[12] || (typeof c[13] === 'number' && dd !== c[13])) {
      malos++;
      if (malos <= 20) Logger.log('Diferencia: %s (fila %s): planilla vence %s / %s días, app calcula %s / %s días', c[1], c[0], c[12], c[13], v, dd);
    }
  });
  Logger.log('Revisados %s clientes activos: %s diferencias.', revisados, malos);
}

function guardarCliente(token, row, nombreOriginal, datos, esperado) {
  const quien = usuario_(token);
  const t0 = Date.now();
  try {
  return conLock_(function () {
    const { ss, tz } = ctx_();
    const h = hojaClientes_(ss);
    const fr = filaYValores_(h, row, nombreOriginal), r = fr.r;
    let renombrados = 0;
    const nuevoNombre = datos.nombre != null ? String(datos.nombre).trim() : nombreOriginal;
    if (!nuevoNombre) throw new Error('El nombre no puede quedar vacío.');
    const cambiaNombre = nuevoNombre !== String(nombreOriginal).trim();
    if (cambiaNombre && nombreExiste_(h, nuevoNombre, r)) {
      throw new Error('Ya hay un cliente llamado "' + nuevoNombre + '". Agregale un emoji o cambiá el nombre.');
    }
    const antes = enc_(fr.fila, r, h.cols, tz);
    verificarSinCambios_(antes, esperado);
    datos = Object.assign({}, datos); delete datos.baja; delete datos.vencio;   // esas dos las maneja solo el servidor
    if (String(datos.fecha || '').trim().toUpperCase() === 'X' && antes[4].toUpperCase() !== 'X') {
      exigir_(quien, 'darDeBaja');
      Object.assign(datos, datosDeBaja_(antes, h, tz));
    }
    if (antes[4].toUpperCase() === 'X' && /^\d{4}-/.test(String(datos.fecha || '')) && h.cols.baja != null) { datos.baja = ''; datos.vencio = ''; }
    if (antes[4].toUpperCase() === 'X' && /^\d{4}-/.test(String(datos.fecha || ''))) {
      validarObligatorios_(Object.assign({ nombre: antes[1], id: antes[2], monto: antes[5], telefono: antes[6], correo: antes[7],
        servicios: antes[8], dispositivos: antes[9], meses: antes[15] }, datos));
    }
    const cambios = escribirCampos_(h, r, datos, tz, fr.fila);
    if (cambiaNombre) renombrados = renombrarEnServicios_(ss, String(nombreOriginal).trim(), nuevoNombre);
    // Velocidad: como al renovar, no se espera a que la planilla recalcule todas sus fórmulas.
    // Se confirma que las celdas cambiadas quedaron escritas y el resto (Perfil, Vencimiento, Días) se calcula acá.
    let despues = CONFIG.RESPUESTA_RAPIDA ? despuesDeCambios_(antes, cambios, tz) : null;
    if (despues) confirmarEscrito_(h, r, cambios, tz);
    else despues = leerCliente_(h, r, tz);
    if (cambios.length) {
      const eraBaja = antes[4].toUpperCase() === 'X', esBaja = despues[4].toUpperCase() === 'X';
      const accion = eraBaja && !esBaja ? 'Reactivó' : !eraBaja && esBaja ? 'Dio de baja' : 'Editó';
      let detalle = detalleVenc_(antes, despues, cambios);
      if (renombrados) detalle += '; nombre cambiado también en ' + renombrados + ' lugar(es) de las hojas de servicio';
      registrar_(ss, quien, accion, despues[1], detalle);
    }
    return { cliente: despues, renombrados: renombrados };
  });
  } finally { console.log('guardarCliente: %s ms', Date.now() - t0); }
}

/**
 * Da de baja como lo hacés a mano: X en Fecha, -N en Monto (días que le esperaste)
 * y "cerrar" al lado de sus lugares, para acordarte de cambiar la contraseña.
 * El nombre queda en su lugar, reservado por si vuelve.
 */
function darDeBaja(token, row, nombre, diasEspera, marcas) {
  const quien = usuario_(token);
  exigir_(quien, 'darDeBaja');
  const t0 = Date.now();
  try {
  olvidarResumen_();
  return conLock_(function () {
    const { ss, tz } = ctx_();
    const h = hojaClientes_(ss);
    const fr = filaYValores_(h, row, nombre), r = fr.r;
    const antes = enc_(fr.fila, r, h.cols, tz);
    const dias = Math.max(0, Math.round(Number(diasEspera) || 0));
    const datos = Object.assign({ fecha: 'X' }, datosDeBaja_(antes, h, tz));
    const mora = moraDe_(antes[8]);
    if (dias > 0) datos.monto = mora ? -mora : -dias;   // Tigo Sports: si le esperaste días queda la mora (-5000)
    const cambios = escribirCampos_(h, r, datos, tz, fr.fila);
    const despues = leerCliente_(h, r, tz);
    const resultado = (marcas || []).map(function (m) {
      try {
        const sh = hojaServicio_(ss, m.hoja);
        const res = marcarCerrar_(sh, Number(m.row), antes[1]);
        res.hoja = m.hoja; res.row = Number(m.row);
        if (!res.cuenta) res.cuenta = 'fila ' + Number(m.row);   // el nombre de la cuenta sale de la misma lectura
        return res;
      } catch (e) {
        return { hoja: m.hoja, row: Number(m.row), ok: false, motivo: String(e.message || e) };
      }
    });
    SpreadsheetApp.flush();
    let detalle = (dias ? 'esperó ' + dias + (dias === 1 ? ' día; ' : ' días; ') : '') + (dias > 0 && mora ? 'mora Gs. ' + mora + '; ' : '') + detalleVenc_(antes, despues, cambios);
    const hechas = resultado.filter(function (x) { return x.ok; });
    if (hechas.length) detalle += '; marcado "cerrar" en ' + hechas.map(function (x) { return x.hoja + ' ' + x.cuenta; }).join(', ');
    registrar_(ss, quien, 'Dio de baja', despues[1], detalle);
    return { cliente: despues, marcas: resultado };
  });
  } finally { console.log('darDeBaja: %s ms', Date.now() - t0); }
}

function nuevoCliente(token, datos) {
  const quien = usuario_(token);
  const t0 = Date.now();
  try {
  return conLock_(function () {
    const { ss, tz } = ctx_();
    const h = hojaClientes_(ss);
    const sh = h.sh;
    const nombre = String(datos.nombre || '').trim();
    validarObligatorios_(datos);
    if (nombreExiste_(h, nombre, -1)) {
      throw new Error('Ya hay un cliente llamado "' + nombre + '". Agregale un emoji o cambiá el nombre.');
    }
    const last = ultimaFila_(sh, h.cols.nombre);
    const r = last + 1;
    if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
    const destino = sh.getRange(r, 1, 1, h.ncol);
    const restaurar = filtroFuera_(sh, r);   // si la fila nueva queda oculta por un filtro, se saca un momento
    try {

    // Formato y fórmulas (Perfil, Vencimiento, Días, ESTADO, SMS) copiados de la fila anterior
    if (last >= 2) {
      const origen = sh.getRange(last, 1, 1, h.ncol);
      origen.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      destino.clearDataValidations();
      const manuales = columnasManuales_(h.cols);
      origen.getFormulasR1C1()[0].forEach(function (f, c) {
        if (f && !manuales[c]) sh.getRange(r, c + 1).setFormulaR1C1(f);
      });
    }

    const completo = Object.assign({
      id: '', fecha: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
      monto: '', correo: '', servicios: '', dispositivos: '', pin: '', pins: '', meses: 1
    }, datos, { nombre: nombre });
    escribirCampos_(h, r, completo, tz);
    } finally { if (restaurar) restaurar(); }
    const c = leerCliente_(h, r, tz);
    registrar_(ss, quien, 'Cargó cliente nuevo', c[1],
      [c[8] || 'sin servicios', 'vence ' + isoDmy_(c[12]), c[3] ? c[3] + ' perfil(es)' : '',
       typeof c[5] === 'number' ? 'Gs. ' + c[5] : '', 'tel. ' + c[6], 'fila ' + r].filter(String).join('; '));
    return c;
  });
  } finally { console.log('nuevoCliente: %s ms', Date.now() - t0); }
}

/* ------------------------------------------------------------------ */
/* Escrituras en hojas de servicio                                     */
/* ------------------------------------------------------------------ */

function asignarLugar(token, hoja, row, col, nombre) {
  const quien = usuario_(token);
  const t0 = Date.now();
  try {
  olvidarResumen_();
  return conLock_(function () {
    const { ss } = ctx_();
    const h = hojaClientes_(ss);
    const n = String(nombre).trim();
    if (!nombreExiste_(h, n, -1)) throw new Error('"' + n + '" no está en BasedeDatos.');
    const sh = hojaServicio_(ss, hoja);
    const cell = sh.getRange(row, col);
    validarCeldaUsuario_(sh, row, col);
    if (cell.getFormula()) throw new Error('Esa celda tiene una fórmula. No la toco.');
    const actual = String(cell.getDisplayValue()).trim();
    if (actual) throw new Error('Ese lugar ya lo ocupa ' + actual + '. Tocá Actualizar.');
    cell.setValue(n);
    SpreadsheetApp.flush();
    const data = cuentaDe_(ss, hoja, row);
    registrar_(ss, quien, 'Asignó lugar', n, hoja + ', ' + cuentaDeFila_(data, row) + ' (fila ' + row + ')');
    return data;
  });
  } finally { console.log('asignarLugar: %s ms', Date.now() - t0); }
}

function liberarLugar(token, hoja, row, col, nombreActual) {
  const quien = usuario_(token);
  exigir_(quien, 'liberarLugar');
  const t0 = Date.now();
  try {
  olvidarResumen_();
  return conLock_(function () {
    const { ss } = ctx_();
    const sh = hojaServicio_(ss, hoja);
    const cell = sh.getRange(row, col);
    validarCeldaUsuario_(sh, row, col);
    if (String(cell.getDisplayValue()).trim() !== String(nombreActual).trim()) {
      throw new Error('Ese lugar cambió en la planilla. Tocá Actualizar.');
    }
    cell.clearContent();
    SpreadsheetApp.flush();
    const data = cuentaDe_(ss, hoja, row);
    registrar_(ss, quien, 'Liberó lugar', String(nombreActual).trim(), hoja + ', ' + cuentaDeFila_(data, row) + ' (fila ' + row + ')');
    return data;
  });
  } finally { console.log('liberarLugar: %s ms', Date.now() - t0); }
}

function cambiarPassCuenta(token, hoja, row, passCol, passAnterior, passNueva, ultimoCambioCol, limpiarCerrar) {
  const quien = usuario_(token);
  exigir_(quien, 'cambiarPass');
  const t0 = Date.now();
  try {
  olvidarResumen_();
  return conLock_(function () {
    const { ss, tz } = ctx_();
    const sh = hojaServicio_(ss, hoja);
    const nueva = String(passNueva || '').trim();
    if (!nueva) throw new Error('La contraseña no puede quedar vacía.');
    const cell = sh.getRange(row, passCol);
    if (cell.getFormula()) throw new Error('La contraseña es una fórmula en la planilla. Cambiala ahí.');
    if (String(cell.getDisplayValue()).trim() !== String(passAnterior || '').trim()) {
      throw new Error('La contraseña cambió en la planilla. Tocá Actualizar.');
    }
    cell.setValue(textoLiteral_(nueva));
    if (ultimoCambioCol) {
      const uc = sh.getRange(row, ultimoCambioCol);
      if (!uc.getFormula()) uc.setValue(Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy') + '\n(último cambio)');
    }
    let quitadas = 0;
    if (limpiarCerrar) {
      const c0 = cuentaDe_(ss, hoja, row).cuentas.filter(function (c) { return c.infoRow === row; })[0];
      if (c0) c0.slots.forEach(function (sl) { if (sl.cerrar) quitadas += quitarCerrar_(sh, sl.row); });
    }
    SpreadsheetApp.flush();
    const data = cuentaDe_(ss, hoja, row);
    const cta = data.cuentas.filter(function (c) { return c.infoRow === row; })[0];
    registrar_(ss, quien, 'Cambió contraseña', '',
      hoja + ', ' + (cta ? [cta.nombre, cta.correo].filter(String).join(' ') : 'fila ' + row) + ': ' + String(passAnterior) + ' → ' + nueva +
      (quitadas ? '; quitó ' + quitadas + ' "cerrar"' : ''));
    return data;
  });
  } finally { console.log('cambiarPassCuenta: %s ms', Date.now() - t0); }
}

/* ------------------------------------------------------------------ */
/* Sesión y registro de cambios                                        */
/* ------------------------------------------------------------------ */

/** Nombre de la propiedad donde se guarda el PIN de cada uno: "Marcelo Benitez" -> PIN_MARCELO_BENITEZ */
function clavePin_(usuario) {
  return 'PIN_' + String(usuario).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');
}
/*
 * Velocidad: las propiedades del proyecto guardan también el recorrido de Pendientes, los pagos usados, etc.
 * (varios cientos de KB). Antes, cada pedido de la app las leía enteras solo para revisar el PIN y la sesión.
 *  - props_(): se leen una sola vez por pedido (antes, hasta 4 o 5 veces en una renovación con pago).
 *  - auth_(): los PIN y el secreto de la sesión quedan en la caché del script 5 minutos.
 *    Si cambiás un PIN en Configuración del proyecto, ejecutá probar() (o esperá 5 minutos) para que valga.
 */
let PROPS_ALL_ = null;
function props_() {
  if (!PROPS_ALL_) PROPS_ALL_ = PropertiesService.getScriptProperties().getProperties();
  return PROPS_ALL_;
}
let AUTH_ = null;
const AUTH_KEY_ = 'auth_v1';
function auth_() {
  if (AUTH_) return AUTH_;
  const cache = CacheService.getScriptCache();
  try { const g = cache.get(AUTH_KEY_); if (g) { AUTH_ = JSON.parse(g); if (AUTH_ && AUTH_.secreto) return AUTH_; } } catch (e) { }
  const p = props_();
  let s = p.SECRETO_SESION;
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('SECRETO_SESION', s);
    p.SECRETO_SESION = s;
  }
  const pins = {};
  Object.keys(CONFIG.USUARIOS || {}).forEach(function (u) { const k = clavePin_(u); if (p[k]) pins[k] = String(p[k]).trim(); });
  AUTH_ = { secreto: s, pins: pins };
  try { cache.put(AUTH_KEY_, JSON.stringify(AUTH_), 300); } catch (e) { }
  return AUTH_;
}
function olvidarAuth_() {
  AUTH_ = null; PROPS_ALL_ = null;
  try { CacheService.getScriptCache().remove(AUTH_KEY_); } catch (e) { }
}
function pinDe_(usuario) {
  const guardado = auth_().pins[clavePin_(usuario)];
  if (guardado) return String(guardado).trim();
  return String((CONFIG.USUARIOS || {})[usuario] || '').trim();
}
function usuariosValidos_() {
  const vistos = {};
  return Object.keys(CONFIG.USUARIOS || {}).filter(function (u) {
    const pin = pinDe_(u);
    if (!/^\d{6}$/.test(pin) || vistos[pin] || /\|/.test(u)) return false;
    vistos[pin] = true;
    return true;
  });
}
function problemasUsuarios_() {
  const vistos = {}, out = [];
  Object.keys(CONFIG.USUARIOS || {}).forEach(function (u) {
    const pin = pinDe_(u);
    if (/\|/.test(u)) out.push(u + ': el nombre no puede tener el carácter |');
    else if (!pin) out.push(u + ': no tiene PIN. Agregá la propiedad ' + clavePin_(u) + ' en Configuración del proyecto. Hasta entonces no puede entrar.');
    else if (!/^\d{6}$/.test(pin)) out.push(u + ': el PIN tiene que tener exactamente 6 números. Hasta corregirlo no puede entrar.');
    else if (vistos[pin]) out.push(u + ': tiene el mismo PIN que ' + vistos[pin] + '. Ponele uno distinto.');
    else vistos[pin] = u;
  });
  return out;
}

/** Aviso por correo cuando la entrada se traba por muchos PIN errados. */
function avisarBloqueo_() {
  try {
    const para = CONFIG.EMAIL_ALERTAS || Session.getEffectiveUser().getEmail();
    if (!para) return;
    const cuando = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    MailApp.sendEmail(para, 'AccountStore: ' + MAX_FALLOS_ + ' intentos de PIN incorrectos',
      'El ' + cuando + ' hubo ' + MAX_FALLOS_ + ' intentos seguidos con un PIN incorrecto en la app.\n' +
      'La entrada quedó trabada 15 minutos.\n\n' +
      'Si no fue nadie de tu equipo: cambiá los PIN en Code.gs, hacé Nueva versión y ejecutá cerrarTodasLasSesiones() desde el editor.\n' +
      'Los intentos quedan anotados en la hoja RegistroApp.');
    registrar_(null, '(sistema)', 'Entrada trabada', '', MAX_FALLOS_ + ' PIN incorrectos seguidos; aviso enviado a ' + para);
  } catch (e) {
    console.warn('No se pudo enviar el aviso: ' + e);
  }
}

/**
 * Ejecutala desde el editor si perdés un celular o sospechás algo raro:
 * corta la sesión en todos los celulares (cada uno vuelve a poner su PIN)
 * y destraba la entrada si estaba trabada por intentos.
 */
function cerrarTodasLasSesiones() {
  PropertiesService.getScriptProperties().deleteProperty('SECRETO_SESION');
  olvidarAuth_();
  CacheService.getScriptCache().remove('pinFallos');
  registrar_(null, '(editor)', 'Cerró todas las sesiones', '', 'todos tienen que volver a entrar con su PIN');
  Logger.log('Listo: se cerraron todas las sesiones. Cada uno tiene que volver a entrar con su PIN.');
}

const PERMISOS_ = { cambiarPass: 'cambiar contraseñas de cuentas', liberarLugar: 'liberar lugares', darDeBaja: 'dar de baja clientes',
  destrabarPago: 'destrabar transferencias ya usadas' };
function permisos_(usuario) {
  const no = (CONFIG.RESTRICCIONES || {})[usuario] || [];
  const out = {};
  Object.keys(PERMISOS_).forEach(function (k) { out[k] = no.indexOf(k) < 0; });
  return out;
}
function exigir_(usuario, accion) {
  if (!permisos_(usuario)[accion]) throw new Error('No tenés permiso para ' + PERMISOS_[accion] + '. Pedíselo a un administrador.');
}

function secreto_() { return auth_().secreto; }
function firma_(texto) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(texto, secreto_()));
}
// Cambia si se cambia el PIN: así una sesión vieja deja de servir
function huellaPin_(usuario) {
  return firma_('pin|' + usuario + '|' + pinDe_(usuario)).slice(0, 12);
}
function firmar_(usuario) {
  const datos = [usuario, Date.now() + CONFIG.DIAS_SESION * 86400000, huellaPin_(usuario)].join('|');
  return Utilities.base64EncodeWebSafe(datos, Utilities.Charset.UTF_8) + '.' + firma_(datos);
}
/** Devuelve el usuario de la sesión, o corta con SESION si no es válida. */
function usuario_(token) {
  const partes = String(token || '').split('.');
  if (partes.length !== 2) throw new Error('SESION');
  let datos;
  try { datos = Utilities.newBlob(Utilities.base64DecodeWebSafe(partes[0])).getDataAsString('UTF-8'); }
  catch (e) { throw new Error('SESION'); }
  if (firma_(datos) !== partes[1]) throw new Error('SESION');
  const d = datos.split('|');
  if (usuariosValidos_().indexOf(d[0]) < 0 || Number(d[1]) < Date.now() || d[2] !== huellaPin_(d[0])) throw new Error('SESION');
  return d[0];
}

/** Anota una línea en la hoja RegistroApp (la crea si no existe). Nunca corta la operación. */
function registrar_(ss, usuario, accion, cliente, detalle) {
  try {
    ss = ss || SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sh = ss.getSheetByName(CONFIG.HOJA_LOG);
    if (!sh) {
      sh = ss.insertSheet(CONFIG.HOJA_LOG);
      sh.getRange(1, 1, 1, 5).setValues([['Fecha', 'Usuario', 'Acción', 'Cliente', 'Detalle']]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(1, 130); sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 180); sh.setColumnWidth(5, 520);
      sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    }
    sh.appendRow([new Date(), usuario, accion, cliente || '', detalle || '']);
  } catch (e) {
    console.warn('No se pudo anotar en el registro: ' + e);
  }
  anotarFechaBaja_(cliente, accion);
}

const ETIQUETAS_ = { nombre: 'Nombre', id: 'ID', perfil: 'Perfil', fecha: 'Activación', monto: 'Monto',
  telefono: 'Teléfono', correo: 'Correo', servicios: 'Servicios', dispositivos: 'Dispositivos',
  pin: 'PIN', pins: 'PINS', meses: 'Meses', baja: 'Baja', vencio: 'Venció' };

function isoDmy_(iso) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : String(iso || '');
}
// "Vencimiento: 19/09/2026 → 19/10/2026; Meses: 31 → 32; Monto: ..."
function detalleVenc_(antes, despues, cambios) {
  const partes = [];
  if (antes[12] !== despues[12]) partes.push('Vencimiento: ' + (isoDmy_(antes[12]) || 'vacío') + ' → ' + (isoDmy_(despues[12]) || 'vacío'));
  cambios.forEach(function (c) { partes.push(ETIQUETAS_[c.campo] + ': ' + (c.antes || 'vacío') + ' → ' + (c.despues || 'vacío')); });
  return partes.join('; ');
}
/**
 * Con varias personas usando la app: si otro renovó o reactivó a este cliente
 * mientras tanto, no se vuelve a sumar encima (evita cobrar dos veces el mismo mes).
 */
function verificarSinCambios_(actual, esperado) {
  if (!esperado) return;
  if (String(actual[4]) !== String(esperado.fecha) || String(actual[15]) !== String(esperado.meses)) {
    throw new Error('CAMBIADO: Otra persona modificó a este cliente hace un momento' +
      (/^\d{4}-/.test(actual[12]) ? ' y ahora vence el ' + isoDmy_(actual[12]) : '') +
      '. Revisá antes de volver a guardar.');
  }
}

/* Campos obligatorios al cargar o reactivar un cliente */
function claveRegla_(s) { return norm_(String(s || '').replace(/-\d+$/, '')).replace(/ /g, ''); }
function aplicaRegla_(lista, servicio) {
  const k = claveRegla_(servicio);
  return (lista || []).some(function (r) { return k.indexOf(claveRegla_(r)) === 0; });
}
function validarObligatorios_(d) {
  const falta = [];
  const servs = String(d.servicios || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x && x !== '-'; });
  const vacio = function (v) { v = String(v == null ? '' : v).trim(); return !v || v === '-'; };
  if (vacio(d.nombre)) falta.push('nombre');
  if (String(d.telefono || '').replace(/\D/g, '').length < 8) falta.push('teléfono');
  if (!servs.length) falta.push('servicios');
  if (vacio(d.id)) falta.push('ID');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.fecha || ''))) falta.push('activación');
  if (!(Number(d.meses) >= 1)) falta.push('meses');
  if (!(Number(String(d.monto == null ? '' : d.monto).replace(/[^\d]/g, '')) > 0)) falta.push('monto');
  const sinDisp = servs.length && servs.every(function (sv) { return aplicaRegla_(CONFIG.SIN_DISPOSITIVO, sv); });
  if (sinDisp) d.dispositivos = '-';
  else if (vacio(d.dispositivos)) falta.push('dispositivos');
  if (servs.some(function (sv) { return aplicaRegla_(CONFIG.CON_CORREO, sv); }) && !/^\S+@\S+\.\S+$/.test(String(d.correo || '').trim())) falta.push('correo');
  if (falta.length) throw new Error('Falta completar: ' + falta.join(', ') + '.');
}

function cuentaDeFila_(data, row) {
  const c = data.cuentas.filter(function (x) { return x.slots.some(function (s) { return s.row === row; }); })[0];
  return c ? (c.nombre || c.plan || 'grupo fila ' + c.header) : 'fila ' + row;
}

/* ------------------------------------------------------------------ */
/* Pendientes: cuentas para cerrar y lugares para liberar              */
/* ------------------------------------------------------------------ */

/**
 * Recorre todas las hojas de servicio y junta:
 *  - lugares marcados "cerrar" (falta cambiar la contraseña de esa cuenta)
 *  - lugares ocupados por dados de baja o por nombres que no están en la base
 */
/**
 * Pendientes: sale del último recorrido guardado (lo hace un proceso automático cada 30 min),
 * así abre al instante. Con forzar = true recorre todo en el momento ("Revisar ahora").
 */
function getPendientes(token, forzar) {
  usuario_(token);
  const t0 = Date.now();
  let a = forzar ? null : leerAnalisis_();
  if (!a) a = recalcularAnalisis_();
  console.log('getPendientes: %s ms%s', Date.now() - t0, forzar ? ' (revisado en el momento)' : '');
  return { hojas: a.hojas, ts: a.resumen.ts };
}

/** Números del resumen del día (del mismo recorrido guardado). */
function getResumen(token) {
  usuario_(token);
  const a = leerAnalisis_() || recalcularAnalisis_();
  return a.resumen;
}

/**
 * Un cliente leído en el momento. La app lo usa si al renovar se cortó la conexión o tardó demasiado:
 * así sabe si la renovación quedó guardada (el servidor la termina aunque la app se haya colgado).
 */
function getCliente(token, row, nombre) {
  usuario_(token);
  const { ss, tz } = ctx_();
  const h = hojaClientes_(ss);
  const fr = filaYValores_(h, row, nombre);
  return enc_(fr.fila, fr.r, h.cols, tz);
}

/** Una cuenta leída en el momento, tal cual se ve en la planilla (para cambiarle la contraseña). */
function getCuenta(token, hoja, row) {
  usuario_(token);
  const { ss } = ctx_();
  return cuentaDe_(ss, hoja, Number(row));
}

/* ---------- Transferencias ---------- */

/** Lo que hay en la hoja TRANSFERENCIAS (rápido: no toca el correo). */
function getTransferencias(token) {
  usuario_(token);
  const { ss, tz } = ctx_();
  return leerTransferencias_(ss, tz);
}

/**
 * Botón "Cargar transferencias": llama a tu script (tu web app) para que lea los correos
 * del banco y agregue las nuevas; después devuelve la hoja actualizada.
 * Si otra persona ya está cargando, espera a que termine y no vuelve a disparar el script.
 */
function cargarTransferencias(token) {
  const quien = usuario_(token);
  const cache = CacheService.getScriptCache(), props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  let nuevas = null, aviso = '', yaCargo = '', filasNuevas = null;
  const enCurso = cache.get('transf_cargando');
  const ultima = Number(props.getProperty('transf_ultima_carga') || 0);
  const t0 = Date.now();
  if (enCurso) {
    // otra persona ya tocó el botón: espero a que termine y no disparo el script de nuevo
    for (let i = 0; i < 45 && cache.get('transf_cargando'); i++) Utilities.sleep(1000);
    yaCargo = enCurso;
  } else if (Date.now() - ultima < 60000) {
    // se cargó hace menos de un minuto: no vale la pena volver a leer el correo
    yaCargo = props.getProperty('transf_ultima_quien') || '';
  } else {
    cache.put('transf_cargando', quien, 120);
    try {
      // CLAVE_TRANSFERENCIAS: la misma clave guardada en las propiedades de los dos proyectos (así nadie más puede usar el link)
      const clave = String(props.getProperty('CLAVE_TRANSFERENCIAS') || '').trim();
      const url = CONFIG.TRANSFERENCIAS_URL + (CONFIG.TRANSFERENCIAS_URL.indexOf('?') >= 0 ? '&' : '?') + 'formato=json' +
        (clave ? '&clave=' + encodeURIComponent(clave) : '');
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      const txt = String(resp.getContentText() || '');
      let j = null;
      try { j = JSON.parse(txt); } catch (e) { }
      if (j && j.ok) { nuevas = Number(j.nuevas) || 0; filasNuevas = j.filas || []; }
      else {
        // script viejo (responde texto): igual sirve, pero hay que releer la hoja
        const m = txt.match(/\((\d+)\s+nueva/i);
        if (m) nuevas = Number(m[1]);
        if (resp.getResponseCode() !== 200 || /^\s*error/i.test(txt)) aviso = 'El script de transferencias respondió: ' + txt.slice(0, 140);
      }
      props.setProperties({ transf_ultima_carga: String(Date.now()), transf_ultima_quien: quien });
    } finally { cache.remove('transf_cargando'); }
  }
  const tScript = Date.now() - t0;
  const r = { nuevas: nuevas, aviso: aviso, ultimaCarga: Number(props.getProperty('transf_ultima_carga') || 0),
              hoy: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), lista: null, filasNuevas: filasNuevas || [] };
  if (filasNuevas == null && !yaCargo) {
    // sin las filas del script: leo la hoja como antes
    const t1 = Date.now();
    const x = ctx_();
    Object.assign(r, leerTransferencias_(x.ss, x.tz));
    console.log('cargarTransferencias: script %s ms, leer hoja %s ms', tScript, Date.now() - t1);
  } else {
    console.log('cargarTransferencias: script %s ms, %s nuevas (sin releer la hoja)', tScript, r.filasNuevas.length);
  }
  r.yaCargo = yaCargo && yaCargo !== quien ? yaCargo : (yaCargo ? 'vos' : '');
  if (!r.lista) Object.assign(r, controlPagos_(r.filasNuevas));
  return r;
}

function leerTransferencias_(ss, tz) {
  const sh = ss.getSheetByName(CONFIG.HOJA_TRANSFERENCIAS);
  const ultimaCarga = Number(PropertiesService.getScriptProperties().getProperty('transf_ultima_carga') || 0);
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (!sh || sh.getLastRow() < 2) return { lista: [], ultimaCarga: ultimaCarga, hoy: hoy };
  const vals = sh.getRange(2, 1, Math.min(sh.getLastRow() - 1, 400), 7).getValues();
  const lista = [];
  vals.forEach(function (v) {
    if (!v[0] && !v[4]) return;
    const f = v[0] instanceof Date ? Utilities.formatDate(v[0], tz, "yyyy-MM-dd'T'HH:mm:ss") : String(v[0]);
    // [fecha, entidad pagadora, cliente pagador, moneda y monto, cuenta donde entró (etiqueta), id, comprobante]
    // (el número de cuenta del pagador no se manda a la app; solo el comprobante, si lo hay)
    const d = String(v[3] || '').trim();
    const extra = /^comprobante/i.test(d) ? d : '';
    // [7] cuenta del pagador, tal como la manda el banco (enmascarada), para el detalle
    lista.push([f, String(v[1] || ''), String(v[2] || ''), String(v[4] || ''), String(v[5] || ''), String(v[6] || ''), extra, extra ? '' : d]);
  });
  return Object.assign({ lista: lista, ultimaCarga: ultimaCarga, hoy: hoy }, controlPagos_(lista));
}

/* ---------- Control de pagos: cada transferencia se usa una sola vez, y quién paga por quién ---------- */
const PAGO_DIAS_USADOS = 90;   // se recuerdan las transferencias usadas de los últimos 90 días
function clavePagador_(s) { return norm_(s); }
function pagosUsados_() { return leerGrande_('pagosUsados') || {}; }
function vinculosPagos_() { return leerGrande_('pagosVinc') || {}; }
/** Para la app: de las transferencias de la lista, cuáles están usadas y a qué clientes se vinculó cada pagador. */
function controlPagos_(lista) {
  const u = pagosUsados_(), v = vinculosPagos_(), usados = {}, vinculos = {};
  (lista || []).forEach(function (f) {
    if (u[f[5]]) usados[f[5]] = u[f[5]];
    const k = clavePagador_(f[2]);
    if (k && v[k]) vinculos[k] = v[k];
  });
  return { usados: usados, vinculos: vinculos };
}
function fechaCortaPago_(ts) { return Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd/MM HH:mm'); }
const METODOS_PAGO = { transferencia: 'Transferencia', billetera: 'Billetera electrónica', pagopar: 'Link de pago (Pagopar)', efectivo: 'Efectivo' };
/** La clave única de un pago: la transferencia de la lista (su ID), o el método + el número de referencia. */
function clavePago_(pago) {
  if (!pago) return '';
  if (pago.id) return String(pago.id);
  const ref = String(pago.ref || '').replace(/[^0-9a-z]/gi, '').toLowerCase();
  return pago.metodo && ref ? 'ref:' + pago.metodo + ':' + ref : '';
}
function textoPago_(pago) {
  if (!pago) return '';
  if (pago.id) return 'transferencia de ' + (pago.pagador || '') + ' (Gs. ' + (pago.monto || '') + ')';
  return (METODOS_PAGO[pago.metodo] || pago.metodo || 'otro medio') + (pago.ref ? ' (ref. ' + pago.ref + ')' : '') +
    (pago.cuenta ? ' en ' + pago.cuenta : '') + (pago.monto ? ', Gs. ' + pago.monto : '');
}
/** Si ese pago (transferencia o número de referencia) ya se usó para otro cliente, corta con un aviso claro. */
function verificarPago_(pago, cliente) {
  const k = clavePago_(pago);
  if (!k) return;
  const u = pagosUsados_()[k];
  if (u && u.cliente !== cliente) {
    throw new Error((pago.id ? 'Esta transferencia' : 'Ese número de referencia') + ' ya se usó para ' + u.cliente + ' (' + u.usuario + ', ' + fechaCortaPago_(u.ts) + '). No se puede usar dos veces.');
  }
}
function marcarPagoUsado_(pago, cliente, quien) {
  if (!pago) return;
  // efectivo sin referencia: se guarda igual (con un id propio) para que aparezca en "Otros pagos"
  const k = clavePago_(pago) || (!pago.id && pago.metodo ? 'efe:' + Date.now() + ':' + Math.round(Math.random() * 1e6) : '');
  if (!k) return;
  const m = pagosUsados_();
  m[k] = { cliente: cliente, usuario: quien, ts: Date.now(), monto: Number(pago.monto) || 0, pagador: String(pago.pagador || ''),
           metodo: pago.id ? 'transferencia' : String(pago.metodo || ''), ref: String(pago.ref || ''),
           cuenta: String(pago.cuenta || ''), manual: !pago.id };
  const lim = Date.now() - PAGO_DIAS_USADOS * 864e5;
  Object.keys(m).forEach(function (k) { if ((m[k].ts || 0) < lim) delete m[k]; });
  guardarGrande_('pagosUsados', m);
  if (pago.id) agregarVinculo_(pago.pagador, cliente);
}

/** "Otros pagos": los cargados a mano al renovar, los más nuevos primero (últimos 90 días). */
function getOtrosPagos(token) {
  usuario_(token);
  const m = pagosUsados_(), out = [];
  Object.keys(m).forEach(function (id) { if (m[id].manual) out.push(Object.assign({ id: id }, m[id])); });
  out.sort(function (a, b) { return b.ts - a.ts; });
  return out;
}

/** Busca un pedido en la hoja PAGOPAR (por número) y dice si ya se usó desde la app. */
function buscarPedidoPagopar(token, numero) {
  usuario_(token);
  const n = String(numero || '').replace(/\D/g, '');
  const usado = n ? (pagosUsados_()['ref:pagopar:' + n] || null) : null;
  if (!n) return { existe: false };
  const { ss } = ctx_();
  const sh = ss.getSheetByName('PAGOPAR');
  if (!sh || sh.getLastRow() < 2) return { existe: false, usado: usado };
  const ancho = Math.min(12, sh.getLastColumn());
  const enc = sh.getRange(1, 1, 1, ancho).getDisplayValues()[0].map(function (x) { return norm_(x); });
  const col = function (re) { for (let i = 0; i < enc.length; i++) if (re.test(enc[i])) return i; return -1; };
  const cel = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(n).findAll()
    .filter(function (r) { return String(r.getDisplayValue()).replace(/\D/g, '') === n; })[0];
  if (!cel) return { existe: false, numero: n, usado: usado };
  const f = sh.getRange(cel.getRow(), 1, 1, ancho).getDisplayValues()[0];
  const v = function (re) { const i = col(re); return i >= 0 ? f[i] : ''; };
  return { existe: true, numero: n, cliente: v(/^cliente/), monto: v(/monto/), estado: v(/^estado/), fecha: v(/^fecha/), producto: v(/^producto/), usado: usado };
}
function agregarVinculo_(pagador, cliente) {
  const k = clavePagador_(pagador);
  if (!k || !cliente || /deposito personal pay/.test(k)) return;   // sin nombre real no se vincula
  const m = vinculosPagos_();
  const l = (m[k] || []).filter(function (x) { return x !== cliente; });
  l.unshift(cliente);
  m[k] = l.slice(0, 6);
  guardarGrande_('pagosVinc', m);
}
/**
 * Vincular una transferencia con un cliente: queda anotado que esa persona paga por ese cliente.
 * Con marcarUsada (el cliente ya se renovó por otro lado), además la transferencia queda usada.
 */
function vincularPago(token, pago, cliente, marcarUsada) {
  const quien = usuario_(token);
  return conLock_(function () {
    const { ss } = ctx_();
    if (marcarUsada) {
      verificarPago_(pago, cliente);
      marcarPagoUsado_(pago, cliente, quien);
      registrar_(ss, quien, 'Usó transferencia', cliente, 'Gs. ' + (pago.monto || '') + ' de ' + (pago.pagador || '') + ' (ya renovado)');
    } else {
      agregarVinculo_(pago.pagador, cliente);
    }
    return controlPagos_([['', '', pago.pagador, '', '', pago.id]]);
  });
}
/** Solo admin y supervisor: deja una transferencia usada otra vez libre (por si fue un error). */
function destrabarPago(token, id) {
  const quien = usuario_(token);
  exigir_(quien, 'destrabarPago');
  return conLock_(function () {
    const { ss } = ctx_();
    const m = pagosUsados_(), u = m[id];
    if (!u) return { ok: true };
    delete m[id];
    guardarGrande_('pagosUsados', m);
    registrar_(ss, quien, 'Destrabó transferencia', u.cliente, 'Gs. ' + u.monto + ' de ' + u.pagador);
    return { ok: true };
  });
}
function quitarVinculo(token, pagador, cliente) {
  usuario_(token);
  return conLock_(function () {
    const m = vinculosPagos_(), k = clavePagador_(pagador);
    if (m[k]) { m[k] = m[k].filter(function (x) { return x !== cliente; }); if (!m[k].length) delete m[k]; guardarGrande_('pagosVinc', m); }
    return { vinculos: m[k] || [] };
  });
}

/* ---------- Ocupación automática: "(LLENO)", "(+N)" y MIEMBRO EXTRA-N ---------- */

function capacidadDe_(hoja) {
  const k = claveHoja_(hoja);
  const cap = CONFIG.CAPACIDAD || {};
  for (const n in cap) if (claveHoja_(n) === k) return Number(cap[n]) || null;
  return null;
}
function vipsGuardados_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('vips') || '[]'); } catch (e) { return []; }
}
// Un lugar está disponible si está vacío o si en la columna G dice LIBRE o CANCELADO
function lugarDisponible_(s) { return !s.usuario || !!s.huerfano || /cancelad/i.test(String(s.fecha || '')); }

/**
 * Anota en cada cuenta (c.ocupacionAuto) su capacidad, perfiles ocupados, disponibles y la marca que corresponde.
 *  - Perfiles de cada fila: columna Perfiles (D); si la hoja no la tiene, 1.
 *  - Cliente repartido en varias filas del mismo servicio: sus perfiles se reparten entre esas filas.
 *  - Clientes VIP: ocupan lugar como todos, salvo en las hojas de VIP_NO_OCUPAN_EN (Disney, ChatGPT).
 *  - Filas marcadas "comparte" (mismo perfil que otro cliente de la cuenta): no suman.
 *  - Cuenta privada en las hojas de PRIVADA_OCUPA_TODO (ej. Crunchyroll-Privada): ocupa todo.
 *  - MIEMBRO EXTRA-N: el número son los lugares ocupados (máximo 2); no lleva marca.
 *  - Cuentas inactivas y bloques de más de MAX_FILAS_OCUPACION filas: sin cálculo.
 */
function calcularOcupacion_(hoja, parsed, vips) {
  if ((CONFIG.SIN_OCUPACION || []).some(function (h) { return claveHoja_(h) === claveHoja_(hoja); })) return parsed;
  const esVip = {};
  const vipNoOcupa = (CONFIG.VIP_NO_OCUPAN_EN || []).some(function (h) { return claveHoja_(h) === claveHoja_(hoja); });
  if (vipNoOcupa) (vips || []).forEach(function (n) { esVip[n] = true; });
  const capServ = capacidadDe_(hoja);
  const filasDe = {};
  parsed.cuentas.forEach(function (c) {
    c.slots.forEach(function (s) { if (s.usuario) (filasDe[s.usuario] = filasDe[s.usuario] || []).push(s); });
  });
  // Cuenta privada (ej. "Crunchyroll-Privada" en la hoja Crunchyroll): ocupa todos los perfiles
  const privadaHoja = (CONFIG.PRIVADA_OCUPA_TODO || []).some(function (h) { return claveHoja_(h) === claveHoja_(hoja); });
  const esPrivada = function (s) {
    return privadaHoja && String(s.servicios || '').split(',').some(function (t) {
      return claveHoja_(String(t).split('-')[0]) === claveHoja_(hoja) && /privad[oa]/i.test(t);
    });
  };
  const perfilDe = new Map();
  const kh = claveHoja_(hoja);
  const desdeServicios = function (s) {   // "Disney-Premium, Office-2" en la hoja OFFICE → 2
    let n = 0;
    String(s.servicios || '').split(',').forEach(function (t) {
      const m = String(t).trim().match(/-(\d+)\s*$/);
      const b = claveHoja_(String(t).replace(/-\d+\s*$/, '').split('-')[0]);
      if (b && (b === kh || b.indexOf(kh) === 0 || kh.indexOf(b) === 0)) n += m ? Number(m[1]) : 1;
    });
    return n;
  };
  Object.keys(filasDe).forEach(function (n) {
    const l = filasDe[n];
    let N = 1;
    l.forEach(function (s) { N = Math.max(N, Number(s.perfil) || (s.perfil ? 0 : desdeServicios(s)) || 0); });
    if (l.length === 1) { perfilDe.set(l[0], N); return; }
    if (N <= l.length) { l.forEach(function (s) { perfilDe.set(s, 1); }); return; }
    const base = Math.floor(N / l.length), extra = N % l.length;
    l.forEach(function (s, i) { perfilDe.set(s, base + (i < extra ? 1 : 0)); });
  });
  parsed.cuentas.forEach(function (c) {
    if (c.inactiva || !c.slots.length) return;
    const ocupa = function (s) { return !lugarDisponible_(s) && !esVip[s.usuario] && !s.comparte; };
    if (/^miembro\s+extra\s*-\s*\d+\s*$/i.test(String(c.plan || ''))) {
      if (c.slots.length !== 2) { c.ocupacionAuto = { tipo: 'raro', motivo: 'dice MIEMBRO EXTRA pero tiene ' + c.slots.length + ' filas' }; return; }
      const ocup = c.slots.filter(ocupa).length;
      c.ocupacionAuto = { tipo: 'miembro', capacidad: 2, usados: ocup, disponibles: Math.max(0, 2 - ocup) };
      return;
    }
    if (c.slots.length > CONFIG.MAX_FILAS_OCUPACION) return;
    const cap = capServ || c.slots.length;
    let usados = 0;
    c.slots.forEach(function (s) { if (ocupa(s)) usados += esPrivada(s) ? cap : (perfilDe.get(s) || 1); });
    const libres = cap - usados;
    c.ocupacionAuto = { tipo: 'marca', capacidad: cap, usados: usados, disponibles: Math.max(0, libres),
                        marca: libres <= 0 ? '(LLENO)' : '(+' + libres + ')' };
  });
  return parsed;
}

/**
 * Qué celdas hay que cambiar en una hoja: la marca de cada cuenta (en la misma columna donde
 * ya las ponés en esa hoja) y el número de las MIEMBRO EXTRA. Solo escribe en celdas vacías o
 * que ya tengan una marca; si hay otro texto, lo informa como conflicto.
 */
function cambiosDeOcupacion_(hoja, parsed, vals) {
  const cambios = [], conflictos = [];
  const usadas = {};
  parsed.cuentas.forEach(function (c) { if (c.ocupacionCol) usadas[c.ocupacionCol] = (usadas[c.ocupacionCol] || 0) + 1; });
  let colHoja = null, max = 0;
  Object.keys(usadas).forEach(function (k) { if (usadas[k] > max) { max = usadas[k]; colHoja = Number(k); } });
  parsed.cuentas.forEach(function (c) {
    const o = c.ocupacionAuto;
    if (!o || !c.infoRow) return;
    const fila = vals[c.infoRow - 1] || [];
    if (o.tipo === 'raro') { conflictos.push({ hoja: hoja, cuenta: c.nombre, row: c.infoRow, col: 1, texto: o.motivo }); return; }
    if (o.tipo === 'miembro') {
      const actual = String(fila[0] || '');
      const nuevo = actual.replace(/-\s*\d+\s*$/, '-' + o.usados);
      if (nuevo !== actual) cambios.push({ hoja: hoja, row: c.infoRow, col: 1, valor: nuevo, antes: actual });
      return;
    }
    const col = c.ocupacionCol || colHoja;
    if (!col) return;   // en esta hoja no usás marcas
    const actual = String(fila[col - 1] || '').trim();
    if (actual && !/^\((lleno|\+\d+)\)$/i.test(actual)) {
      conflictos.push({ hoja: hoja, cuenta: c.nombre, row: c.infoRow, col: col, texto: actual });
      return;
    }
    if (actual.toUpperCase() !== o.marca.toUpperCase()) cambios.push({ hoja: hoja, row: c.infoRow, col: col, valor: o.marca, antes: actual });
  });
  return { cambios: cambios, conflictos: conflictos };
}

/** Escribe los cambios en un solo pedido, sin pisar nunca una celda que tenga fórmula. */
function aplicarOcupacion_(ss, cambios) {
  if (!cambios || !cambios.length) return 0;
  const a1 = function (c) { return rangoHoja_(c.hoja) + '!' + letraCol_(c.col) + c.row; };
  let ok = cambios;
  try {
    if (hayApiSheets_()) {
      const f = Sheets.Spreadsheets.Values.batchGet(CONFIG.SPREADSHEET_ID, { ranges: cambios.map(a1), valueRenderOption: 'FORMULA' });
      ok = cambios.filter(function (c, i) {
        const v = f.valueRanges[i] && f.valueRanges[i].values && f.valueRanges[i].values[0] ? String(f.valueRanges[i].values[0][0]) : '';
        return v.charAt(0) !== '=';
      });
      if (ok.length) Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'RAW',
        data: ok.map(function (c) { return { range: a1(c), values: [[c.valor]] }; }) }, CONFIG.SPREADSHEET_ID);
    } else {
      ok = cambios.filter(function (c) {
        const r = ss.getSheetByName(c.hoja).getRange(c.row, c.col);
        if (r.getFormula()) return false;
        r.setValue(c.valor);
        return true;
      });
    }
  } catch (e) { console.warn('No se pudo actualizar la ocupación: ' + e); return 0; }
  return ok.length;
}
function letraCol_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/**
 * Ejecutala desde el editor: muestra en el registro qué marcas cambiaría la ocupación automática,
 * sin escribir nada. Cuando estés conforme, poné OCUPACION_AUTOMATICA: true y hacé Nueva versión.
 */
function revisarOcupacion() {
  const { ss } = ctx_();
  const r = analizarServicios_(ss);
  const oc = r.ocupacion || { cambios: [], conflictos: [] };
  Logger.log('Cambiaría ' + oc.cambios.length + ' celdas:');
  oc.cambios.forEach(function (c) { Logger.log('  ' + c.hoja + ' ' + letraCol_(c.col) + Math.round(c.row) + ': ' + (c.antes || '(vacío)') + ' → ' + c.valor); });
  Logger.log('Para revisar a mano (no se tocan): ' + oc.conflictos.length);
  oc.conflictos.forEach(function (c) { Logger.log('  ' + c.hoja + ' ' + c.cuenta + ', fila ' + Math.round(c.row) + ': "' + c.texto + '"'); });
  Logger.log(CONFIG.OCUPACION_AUTOMATICA ? 'La ocupación automática está ACTIVADA.' : 'La ocupación automática está APAGADA (OCUPACION_AUTOMATICA: false).');
}

/** Actualiza la ocupación de las hojas indicadas (la app lo llama por detrás después de asignar, liberar o dar de baja). */
function actualizarOcupacion(token, hojas) {
  usuario_(token);
  const t0 = Date.now();
  const { ss } = ctx_();
  const lista = (Array.isArray(hojas) ? hojas : [hojas]).filter(Boolean);
  const datos = leerServiciosRapido_(ss, lista);
  const vips = vipsGuardados_();
  let cambios = [], conflictos = [];
  lista.forEach(function (h) {
    if (!datos[h] || !datos[h].length) return;
    const p = calcularOcupacion_(h, parseServicio_(datos[h], h), vips);
    const x = cambiosDeOcupacion_(h, p, datos[h]);
    cambios = cambios.concat(x.cambios); conflictos = conflictos.concat(x.conflictos);
  });
  const n = CONFIG.OCUPACION_AUTOMATICA ? aplicarOcupacion_(ss, cambios) : 0;
  console.log('actualizarOcupacion (%s): %s marcas cambiadas, %s ms', lista.join(', '), n, Date.now() - t0);
  return { cambiadas: n, conflictos: conflictos };
}

/* ---------- Recorrido automático ---------- */

/** Ejecutala UNA VEZ desde el editor: deja el recorrido automático cada 30 minutos (en el HORARIO_AUTOMATICO). */
function activarPrecalculo() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'precalcularPendientes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('precalcularPendientes').timeBased().everyMinutes(30).create();
  const t0 = Date.now();
  recalcularAnalisis_();
  Logger.log('Listo: el recorrido automático quedó activo cada 30 minutos. Este primero tardó %s segundos.', Math.round((Date.now() - t0) / 1000));
}
/** Por si alguna vez querés apagarlo. */
function desactivarPrecalculo() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'precalcularPendientes') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Recorrido automático apagado.');
}
/** Lo llama el activador cada 30 minutos. Fuera del horario de trabajo no hace nada, para ahorrar cupo. */
function precalcularPendientes() {
  const { ss, tz } = ctx_();
  const hora = Number(Utilities.formatDate(new Date(), tz, 'H'));
  const desde = (CONFIG.HORARIO_AUTOMATICO || [9, 24])[0], hasta = (CONFIG.HORARIO_AUTOMATICO || [9, 24])[1];
  if (hora < desde || hora >= hasta) return;
  const t0 = Date.now();
  recalcularAnalisis_(ss);
  console.log('precalcularPendientes: %s ms', Date.now() - t0);
}
function recalcularAnalisis_(ss) {
  ss = ss || ctx_().ss;
  const r = analizarServicios_(ss);
  // ocupación automática: escribe solo las marcas que cambiaron
  const oc = r.ocupacion || { cambios: [], conflictos: [] };
  const to0 = Date.now();
  const n = CONFIG.OCUPACION_AUTOMATICA ? aplicarOcupacion_(ss, oc.cambios) : 0;
  console.log('Ocupación: %s cambios para escribir, %s escritos, %s ms', oc.cambios.length, n, Date.now() - to0);
  if (n || oc.conflictos.length) console.log('Ocupación: %s celdas actualizadas, %s para revisar', n, oc.conflictos.length);
  r.resumen.marcasCambiadas = n;
  r.conflictosOcupacion = oc.conflictos.slice(0, 50);
  delete r.ocupacion;
  const tg0 = Date.now();
  guardarAnalisis_(r);
  console.log('Guardar el resultado: %s ms', Date.now() - tg0);
  return r;
}

/**
 * Ejecutala desde el editor para ver cuánto tarda cada paso de "Revisar ahora"
 * (el detalle aparece abajo, en el registro de ejecución). Hace lo mismo que el botón.
 */
function medirRecorrido() {
  const t0 = Date.now();
  const { ss } = ctx_();
  console.log('Abrir la planilla: %s ms', Date.now() - t0);
  recalcularAnalisis_(ss);
  console.log('TOTAL: %s ms', Date.now() - t0);
}
// Se guarda en las propiedades del proyecto, en partes de 3.000 caracteres (no vence)
function guardarAnalisis_(r) { guardarGrande_('analisis', r); }
function leerAnalisis_() { return leerGrande_('analisis'); }
/** Guarda un objeto en las propiedades del proyecto, en partes de 3.000 caracteres (no vence). */
function guardarGrande_(prefijo, obj) {
  try {
    const props = PropertiesService.getScriptProperties();
    const txt = JSON.stringify(obj), partes = {};
    const T = 3000;   // 3.000 caracteres x 3 bytes como máximo = 9.000 bytes, debajo del límite de Google
    const n = Math.ceil(txt.length / T);
    for (let i = 0; i < n; i++) partes[prefijo + '_' + i] = txt.slice(i * T, (i + 1) * T);
    const antes = Number(props.getProperty(prefijo + '_n') || 0);
    partes[prefijo + '_n'] = String(n);
    props.setProperties(partes);
    for (let i = n; i < antes; i++) props.deleteProperty(prefijo + '_' + i);
    // la copia en memoria de este pedido queda al día (props_)
    if (PROPS_ALL_) {
      Object.keys(partes).forEach(function (k) { PROPS_ALL_[k] = partes[k]; });
      for (let i = n; i < antes; i++) delete PROPS_ALL_[prefijo + '_' + i];
    }
  } catch (e) { console.warn('No se pudo guardar ' + prefijo + ': ' + e); }
}
function leerGrande_(prefijo) {
  try {
    const props = props_();
    const n = Number(props[prefijo + '_n'] || 0);
    if (!n) return null;
    let txt = '';
    for (let i = 0; i < n; i++) { if (props[prefijo + '_' + i] == null) return null; txt += props[prefijo + '_' + i]; }
    return JSON.parse(txt);
  } catch (e) { return null; }
}

/* ---------- Fechas de baja (para "hace N días" en Pendientes), sin leer RegistroApp ---------- */
function anotarFechaBaja_(cliente, accion) {
  if (accion !== 'Dio de baja' && accion !== 'Reactivó') return;
  try {
    const k = String(cliente || '').trim().toLowerCase();
    if (!k) return;
    const m = leerGrande_('bajasF') || {};
    if (accion === 'Dio de baja') m[k] = Date.now(); else delete m[k];
    const lim = Date.now() - 120 * 864e5;   // se guardan las de los últimos 120 días
    Object.keys(m).forEach(function (x) { if (m[x] < lim) delete m[x]; });
    guardarGrande_('bajasF', m);
  } catch (e) { console.warn('No se pudo anotar la fecha de baja: ' + e); }
}
/** Las fechas de baja guardadas; la primera vez las arma una sola vez desde RegistroApp. */
function fechasBaja_(ss) {
  const props = PropertiesService.getScriptProperties();
  let m = leerGrande_('bajasF') || {};
  if (props.getProperty('bajasF_listo')) return m;
  const desdeLog = {};
  const log = ss.getSheetByName(CONFIG.HOJA_LOG);
  if (log && log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 4).getValues().forEach(function (r) {
    const k = String(r[3]).trim().toLowerCase();
    if (r[2] === 'Dio de baja' && r[0] instanceof Date) desdeLog[k] = r[0].getTime();
    if (r[2] === 'Reactivó') delete desdeLog[k];
  });
  Object.keys(m).forEach(function (k) { desdeLog[k] = m[k]; });   // lo anotado después manda
  const lim = Date.now() - 120 * 864e5;
  Object.keys(desdeLog).forEach(function (x) { if (desdeLog[x] < lim) delete desdeLog[x]; });
  guardarGrande_('bajasF', desdeLog);
  props.setProperty('bajasF_listo', '1');
  return desdeLog;
}

/** Columnas de BasedeDatos que usa el recorrido (se guardan 6 horas para no leer el encabezado cada vez). */
function colsClientes_(ss) {
  const cache = CacheService.getScriptCache();
  const g = cache.get('colsClientes_v1');
  if (g) return JSON.parse(g);
  const h = hojaClientes_(ss);
  const c = { nombre: h.cols.nombre, fecha: h.cols.fecha, monto: h.cols.monto, servicios: h.cols.servicios };
  cache.put('colsClientes_v1', JSON.stringify(c), 21600);
  return c;
}

function olvidarResumen_() { }

/** "Entró a la app": una sola línea por persona y por día en RegistroApp. */
function registrarIngreso_(ss, usuario, tz) {
  try {
    const props = PropertiesService.getScriptProperties();
    const clave = 'ingreso_' + usuario, hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    if (props.getProperty(clave) === hoy) return;
    props.setProperty(clave, hoy);
    registrar_(ss, usuario, 'Entró a la app', '', '');
  } catch (e) { console.warn('No se pudo anotar el ingreso: ' + e); }
}

/**
 * Recorre todas las hojas de servicio una sola vez y arma:
 *  - hojas: lo pendiente (lugares con "cerrar", dados de baja, nombres que no están en la base)
 *  - resumen: lugares libres en cuentas activas, cuentas para cerrar, lugares para liberar
 */
function analizarServicios_(ss) {
  const vips = [];
  const t0 = Date.now();
  const lista = hojasServicio_(ss, false);
  const tLista = Date.now() - t0;
  // En el mismo pedido: las hojas de servicio (A:L) y de BasedeDatos solo nombre, fecha, monto y servicios
  const tc0 = Date.now();
  const cb = colsClientes_(ss);
  const tCols = Date.now() - tc0;
  const campos = ['nombre', 'fecha', 'monto', 'servicios'].filter(function (k) { return cb[k] != null; });
  const hb = "'" + CONFIG.HOJA_CLIENTES.replace(/'/g, "''") + "'!";
  const datos = leerServiciosRapido_(ss, lista, campos.map(function (k) { const L = letraCol_(cb[k] + 1); return hb + L + ':' + L; }));
  const base = {};
  const sumar = function (nom, fecha, monto, serv) {
    const n = String(nom == null ? '' : nom).trim();
    if (!n) return;
    base[n.toLowerCase()] = { baja: String(fecha).trim().toUpperCase() === 'X', monto: monto };
    if (/(^|,)\s*vip\s*(,|$)/i.test(String(serv || ''))) vips.push(n);
  };
  if (datos._extras) {
    const col = {}; campos.forEach(function (k, i) { col[k] = datos._extras[i] || []; });
    for (let i = 1; i < col.nombre.length; i++) sumar(col.nombre[i], (col.fecha || [])[i], (col.monto || [])[i], (col.servicios || [])[i]);
  } else {
    const h = hojaClientes_(ss);
    const last = ultimaFila_(h.sh, h.cols.nombre);
    if (last > 1) h.sh.getRange(2, 1, last - 1, h.ncol).getValues().forEach(function (v) {
      sumar(v[h.cols.nombre], v[h.cols.fecha], v[h.cols.monto], h.cols.servicios != null ? v[h.cols.servicios] : '');
    });
  }
  // cuándo se dio de baja (lo que se hizo desde la app queda anotado)
  const tf0 = Date.now();
  const fb = fechasBaja_(ss), bajaEl = {};
  const tFechas = Date.now() - tf0;
  Object.keys(fb).forEach(function (k) { bajaEl[k] = new Date(fb[k]); });
  const ahora = Date.now();
  const hojas = [];
  const resumen = { libres: 0, cuentas: 0, cerrar: 0, liberar: 0, ts: ahora };
  const tLeer = Date.now() - t0;
  const ocupacion = { cambios: [], conflictos: [] };
  resumen.disponibles = 0;
  lista.forEach(function (hoja) {
    const cuentas = [];
    if (!datos[hoja] || !datos[hoja].length) return;
    const parsed = calcularOcupacion_(hoja, parseServicio_(datos[hoja], hoja), vips);
    const oc = cambiosDeOcupacion_(hoja, parsed, datos[hoja]);
    ocupacion.cambios = ocupacion.cambios.concat(oc.cambios);
    ocupacion.conflictos = ocupacion.conflictos.concat(oc.conflictos);
    parsed.cuentas.forEach(function (c) {
      if (c.ocupacionAuto && !c.inactiva && typeof c.ocupacionAuto.disponibles === 'number') resumen.disponibles += c.ocupacionAuto.disponibles;
      if (c.inactiva) return;   // cuentas para reactivar: no cuentan
      resumen.cuentas++;
      const slots = [];
      c.slots.forEach(function (sl) {
        if (sl.libre) resumen.libres++;
        const info = sl.usuario ? base[sl.usuario.toLowerCase()] : null;
        const esBaja = !!(info && info.baja), noEsta = !!sl.usuario && !info;
        if (!sl.libre && (esBaja || noEsta)) resumen.liberar++;
        if (!sl.cerrar && !esBaja && !noEsta) return;
        const p = { row: sl.row, usuario: sl.usuario || '', libre: !!sl.libre, cerrar: !!sl.cerrar, baja: esBaja, noEsta: noEsta };
        if (esBaja && typeof info.monto === 'number' && info.monto < 0 && info.monto >= -60) p.espero = -info.monto;
        const f = sl.usuario && bajaEl[sl.usuario.toLowerCase()];
        if (esBaja && f) p.hace = Math.floor((ahora - f.getTime()) / 86400000);
        slots.push(p);
      });
      if (c.slots.some(function (sl) { return sl.cerrar; })) resumen.cerrar++;
      // (sin el texto del aviso: Pendientes no lo muestra y al cambiar la contraseña se lee en el momento)
      if (slots.length) cuentas.push({ header: c.header, nombre: c.nombre, plan: c.plan, correo: c.correo, pass: c.pass,
        passCol: c.passCol, infoRow: c.infoRow, ultimoCambioCol: c.ultimoCambioCol, usuarioCol: c.usuarioCol, slots: slots });
    });
    if (cuentas.length) hojas.push({ hoja: hoja, cuentas: cuentas });
  });
  console.log('Recorrido: lista de hojas %s ms, columnas %s ms, leer %s ms (%s), fechas de baja %s ms, total %s ms',
    tLista, tCols, tLeer - tLista - tCols, datos._modo, tFechas, Date.now() - t0);
  try { PropertiesService.getScriptProperties().setProperty('vips', JSON.stringify(vips)); } catch (e) { }
  return { hojas: hojas, resumen: resumen, ocupacion: ocupacion };
}

/** Libera varios lugares de una hoja de una vez (cada uno se verifica antes de borrar). */
function liberarVarios(token, hoja, lugares) {
  const quien = usuario_(token);
  exigir_(quien, 'liberarLugar');
  const t0 = Date.now();
  try {
  olvidarResumen_();
  return conLock_(function () {
    const { ss } = ctx_();
    const sh = hojaServicio_(ss, hoja);
    const hechos = [], saltados = [], filas = [];
    (lugares || []).forEach(function (l) {
      const row = Number(l.row), col = Number(l.col) || 1;
      try { validarCeldaUsuario_(sh, row, col); } catch (e) { saltados.push(l.nombre); return; }
      const cell = sh.getRange(row, col);
      if (cell.getFormula() || String(cell.getDisplayValue()).trim() !== String(l.nombre).trim()) { saltados.push(l.nombre); return; }
      cell.clearContent();
      filas.push(row); hechos.push(String(l.nombre).trim());
    });
    SpreadsheetApp.flush();
    const data = servicio_(ss, hoja);
    if (hechos.length) registrar_(ss, quien, hechos.length === 1 ? 'Liberó lugar' : 'Liberó lugares', '',
      hoja + ': ' + hechos.map(function (h, k) { return h + ' (' + cuentaDeFila_(data, filas[k]) + ')'; }).join(', '));
    return { liberados: hechos.length, saltados: saltados, servicio: data };
  });
  } finally { console.log('liberarVarios: %s ms', Date.now() - t0); }
}

// Encabezados que son datos del lugar; la marca va en la primera columna libre después de ellos
const RX_DATOS_ = /^(usuario|chat|id|perfil|perfiles|numero|dispositivo|dispositivos|tipo de dispositivo|fecha|dias restantes|monto|precio|servicio|servicios|pin|pins|correo|correos)( |$)/;

/**
 * Columna donde va "cerrar" para un lugar: la primera sin fórmula ni título después de los
 * datos, y que en todo el bloque esté vacía o solo tenga marcas (cerrar, *, -).
 * Así nunca pisa una columna de notas o de contraseñas.
 */
function colMarca_(sh, row) {
  return ventanaMarca_(sh, row).col;
}
/** Lee una sola vez las filas alrededor del lugar (valores y fórmulas) para todo lo que necesita marcar "cerrar". */
function ventanaMarca_(sh, row) {
  const n = sh.getLastColumn();
  const desde = Math.max(1, row - 40), hasta = Math.min(sh.getLastRow(), row + 40);
  if (hasta < row) return { vals: [], forms: [], i: -1, h: -1, col: null };
  const rango = sh.getRange(desde, 1, hasta - desde + 1, n);
  const vals = rango.getDisplayValues();
  const forms = rango.getFormulas();
  const i = row - desde;
  const out = { vals: vals, forms: forms, i: i, h: -1, col: null };
  let h = i - 1;
  while (h >= 0 && norm_(vals[h][0]) !== 'usuario') h--;
  out.h = h;
  out.col = colMarcaEn_(vals, forms, i, h, n);
  return out;
}
function colMarcaEn_(vals, forms, i, h, n) {
  while (h >= 0 && norm_(vals[h][0]) !== 'usuario') h--;
  if (h < 0) return null;
  const filas = [];
  for (let r = h + 1; r < vals.length; r++) {
    const a = norm_(vals[r][0]);
    if (a === 'usuario' || /^cuenta\b/.test(a) || !forms[r].slice(1).some(String)) break;
    filas.push(r);
  }
  if (filas.indexOf(i) < 0) return null;
  let ult = 0;
  vals[h].forEach(function (t, c) { if (RX_DATOS_.test(norm_(t))) ult = Math.max(ult, c); });
  const libre = function (v) { return /^\s*(cerrar|\*|-)?\s*$/i.test(String(v)); };
  for (let c = ult + 1; c < n + 3; c++) {
    if (c < n && (String(vals[h][c]).trim() || forms[h][c])) continue;
    if (filas.every(function (r) { return c >= n || (!forms[r][c] && libre(vals[r][c])); })) return c + 1;
  }
  return null;
}
/**
 * Marca "cerrar" al lado del lugar. Velocidad: hace UNA lectura de la zona (antes eran 7 u 8 lecturas
 * por lugar, más otra para saber el nombre de la cuenta). Devuelve { ok, motivo?, cuenta }.
 */
function marcarCerrar_(sh, row, nombre) {
  if (row < 2) throw new Error('Esa celda no parece un lugar de la columna Usuario.');
  const w = ventanaMarca_(sh, row);
  const A = function (k) { return k >= 0 && w.vals[k] ? String(w.vals[k][0]).trim() : ''; };
  // misma regla que validarCeldaUsuario_: "Usuario" en la columna A, como mucho 15 filas arriba
  if (w.h < 0 || w.i - w.h > 15) throw new Error('Esa celda no parece un lugar de la columna Usuario.');
  const esCuenta = function (s) { return /^cuenta\b/.test(norm_(s)); };
  const cuenta = esCuenta(A(w.h - 1)) ? A(w.h - 1) : esCuenta(A(w.h - 2)) ? A(w.h - 2) : (A(w.h - 1) || 'grupo fila ' + (row - w.i + w.h));
  if (A(w.i) !== String(nombre).trim()) return { ok: false, motivo: 'ese lugar ya no es de este cliente', cuenta: cuenta };
  const c = w.col;
  if (!c) return { ok: false, motivo: 'no encontré una columna libre para anotarlo', cuenta: cuenta };
  const actual = String((w.vals[w.i] || [])[c - 1] == null ? '' : w.vals[w.i][c - 1]).trim();
  if (actual && !/^cerrar$/i.test(actual)) return { ok: false, motivo: 'la celda ya tiene "' + actual + '"', cuenta: cuenta };
  sh.getRange(row, c).setValue('cerrar');
  return { ok: true, cuenta: cuenta };
}
function quitarCerrar_(sh, row) {
  const n = sh.getLastColumn();
  const vals = sh.getRange(row, 1, 1, n).getDisplayValues()[0];
  const forms = sh.getRange(row, 1, 1, n).getFormulas()[0];
  let k = 0;
  vals.forEach(function (v, c) {
    if (c > 0 && !forms[c] && /^\s*cerrar\s*$/i.test(v)) { sh.getRange(row, c + 1).clearContent(); k++; }
  });
  return k;
}

/* ------------------------------------------------------------------ */
/* Lectura de hojas de servicio                                        */
/* ------------------------------------------------------------------ */

/**
 * Lectura rápida para el recorrido de Pendientes: solo columnas A a L (ahí están el usuario,
 * los datos de la cuenta y las marcas "cerrar"; de M en adelante van los textos largos).
 * Con el servicio "Google Sheets API" activado lee las 35 hojas en UN solo pedido;
 * si no está activado, lee hoja por hoja (más lento).
 */
const COLS_RECORRIDO = 12;   // A..L
function leerServiciosRapido_(ss, hojas, extras) {
  const conv = function (v) { if (v instanceof Date) return 'fecha'; return v == null ? '' : String(v); };
  const armar = function (fila) { const f = fila.map(conv); while (f.length < COLS_RECORRIDO) f.push(''); return f; };
  const out = {};
  if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
    try {
      const ex = extras || [];
      const r = Sheets.Spreadsheets.Values.batchGet(CONFIG.SPREADSHEET_ID, {
        ranges: hojas.map(function (h) { return "'" + h.replace(/'/g, "''") + "'!A:L"; }).concat(ex),
        valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'FORMATTED_STRING', majorDimension: 'ROWS'
      });
      (r.valueRanges || []).forEach(function (vr, i) {
        if (i < hojas.length) out[hojas[i]] = (vr.values || []).map(armar);
        else (out._extras = out._extras || [])[i - hojas.length] = (vr.values || []).map(function (f) { return f[0] == null ? '' : f[0]; });
      });
      out._modo = 'un pedido';
      return out;
    } catch (e) { console.warn('La lectura en un solo pedido falló, sigo hoja por hoja: ' + e); }
  }
  hojas.forEach(function (h) {
    const sh = hojaServicio_(ss, h);
    const n = sh.getLastRow();
    out[h] = n < 1 ? [] : sh.getRange(1, 1, n, Math.min(COLS_RECORRIDO, sh.getLastColumn())).getValues().map(armar);
  });
  out._modo = 'hoja por hoja';
  return out;
}
function servicio_(ss, hoja) {
  const sh = hojaServicio_(ss, hoja);
  return parseServicio_(sh.getDataRange().getDisplayValues(), hoja);
}

/**
 * Lee solo la cuenta que contiene esa fila (unas 50 filas en vez de toda la hoja).
 * Devuelve { hoja, parcial: true, cuentas: [esa cuenta] }. Si el bloque no entra
 * en la ventana, lee la hoja entera como siempre.
 */
function cuentaDe_(ss, hoja, row) {
  const sh = hojaServicio_(ss, hoja);
  const n = sh.getLastColumn(), last = sh.getLastRow();
  const desde = Math.max(1, row - 25), hasta = Math.min(last, row + 25);
  const top = sh.getRange(1, 1, Math.min(3, last), n).getDisplayValues();
  const ventana = sh.getRange(desde, 1, hasta - desde + 1, n).getDisplayValues();
  const vacia = new Array(n).fill('');
  const vals = new Array(hasta);
  for (let i = 0; i < hasta; i++) vals[i] = vacia;
  top.forEach(function (r, i) { vals[i] = r; });
  ventana.forEach(function (r, i) { vals[desde - 1 + i] = r; });
  const c = parseServicio_(vals, hoja).cuentas.filter(function (x) {
    return x.infoRow === row || x.header === row || x.slots.some(function (s) { return s.row === row; });
  })[0];
  const completa = c && c.header - 2 >= desde &&
    (!c.slots.length || c.slots[c.slots.length - 1].row < hasta || hasta === last);
  if (!completa) { const r = servicio_(ss, hoja); try { calcularOcupacion_(hoja, r, vipsGuardados_()); } catch (e) { } return r; }
  const r = { hoja: hoja, parcial: true, cuentas: [c] };
  try { calcularOcupacion_(hoja, r, vipsGuardados_()); } catch (e) { }   // disponibles de esa cuenta, para la app
  return r;
}

/**
 * Convierte una hoja de servicio en cuentas con sus lugares.
 * Estructura esperada por bloque:
 *   Cuenta N            (fila opcional)
 *   plan | correo | contraseña | último cambio | ciclo | ocupación
 *   Usuario | ... | DIAS RESTANTES | ...   (encabezado de lugares)
 *   lugares...            (hasta una fila vacía)
 */
function parseServicio_(vals, hoja) {
  const ES_ERROR = /^#(N\/A|VALUE!|REF!|DIV\/0!|NAME\?|ERROR!|NUM!|NULL!)/;
  const limpio = function (s) { s = String(s == null ? '' : s).trim(); return (ES_ERROR.test(s) || s === 'ERROR') ? '' : s; };
  const esCuenta = function (s) { return /^cuenta\b/.test(norm_(s)); };
  const vacia = function (row, hasta) {
    for (let c = 0; c < Math.min(row.length, hasta); c++) if (String(row[c]).trim() !== '') return false;
    return true;
  };

  // Etiquetas de las primeras filas (number, correo, pass, dias, PIN, DATA, SUB)
  const top = {};
  for (let r = 0; r < Math.min(3, vals.length); r++) {
    vals[r].forEach(function (v, c) {
      const n = norm_(v);
      if (['correo', 'pass', 'dias', 'pin', 'data', 'sub'].indexOf(n) >= 0 && top[n] == null) top[n] = c;
    });
  }
  const anchoUtil = top.sub != null ? top.sub : (vals[0] ? vals[0].length : 0);

  const cuentas = [];
  for (let h = 0; h < vals.length; h++) {
    const head = vals[h];
    if (norm_(head[0]) !== 'usuario') continue;
    if (!head.some(function (x) { return norm_(x) === 'dias restantes'; })) continue;

    // Columnas de los lugares según el encabezado
    const col = {};
    head.forEach(function (v, c) {
      const n = norm_(v);
      let k = null;
      if (n === 'usuario') k = 'usuario';
      else if (n === 'chat') k = 'chat';
      else if (n === 'id') k = 'id';
      else if (n.indexOf('perfil') === 0) k = 'perfil';
      else if (n.indexOf('numero') >= 0) k = 'numero';
      else if (n.indexOf('dispositivo') >= 0) k = 'dispositivo';
      else if (n === 'fecha') k = 'fecha';
      else if (n === 'dias restantes') k = 'dias';
      else if (n === 'monto' || n === 'precio') k = 'monto';
      else if (n.indexOf('servicio') === 0) k = 'servicios';
      else if (n === 'pin') k = 'pin';
      else if (n.indexOf('correo') === 0) k = 'correo';
      if (k && col[k] == null) col[k] = c;
    });

    // Filas de la cuenta
    let cuentaRow = null, infoRow = null;
    if (h - 1 >= 0 && esCuenta(vals[h - 1][0])) cuentaRow = h - 1;
    else {
      if (h - 1 >= 0 && !vacia(vals[h - 1], anchoUtil) && !vals[h - 1].some(function (x) { return ES_ERROR.test(String(x)); }) && norm_(vals[h - 1][0]) !== 'usuario') infoRow = h - 1;
      if (h - 2 >= 0 && esCuenta(vals[h - 2][0])) cuentaRow = h - 2;
    }

    const cuenta = {
      header: h + 1,
      nombre: cuentaRow != null ? limpio(vals[cuentaRow][0]) : '',
      plan: '', infoRow: infoRow != null ? infoRow + 1 : null,
      correo: '', correoCol: null, pass: '', passCol: null,
      ultimoCambio: '', ultimoCambioCol: null, ciclo: '', ocupacion: '', notas: [],
      aviso: '', usuarioCol: col.usuario + 1, slots: []
    };

    if (infoRow != null) {
      const info = vals[infoRow];
      const usadas = {};
      cuenta.plan = limpio(info[0]); usadas[0] = true;
      // correo
      let cc = (top.correo != null && limpio(info[top.correo])) ? top.correo : null;
      if (cc == null) for (let c = 1; c < info.length; c++) if (/\S+@\S+/.test(info[c])) { cc = c; break; }
      if (cc != null) { cuenta.correo = limpio(info[cc]); cuenta.correoCol = cc + 1; usadas[cc] = true; }
      // contraseña
      let pc = (top.pass != null && limpio(info[top.pass]) && top.pass !== cc) ? top.pass : null;
      if (pc == null && cc != null && limpio(info[cc + 1]) && !/cambio/i.test(info[cc + 1])) pc = cc + 1;
      if (pc != null) { cuenta.pass = limpio(info[pc]); cuenta.passCol = pc + 1; usadas[pc] = true; }
      info.forEach(function (v, c) {
        const t = limpio(v);
        if (!t || usadas[c] || c === top.sub) return;
        const n = norm_(t);
        if (n.indexOf('ultimo cambio') >= 0) {
          cuenta.ultimoCambio = t.split(/\n|\(/)[0].trim(); cuenta.ultimoCambioCol = c + 1;
        } else if (/del mes|reactivar|cancelar/.test(n) && !cuenta.ciclo) {
          cuenta.ciclo = t;
        } else if (/^\(.*\)$/.test(t) && !cuenta.ocupacion) {
          cuenta.ocupacion = t.replace(/^\(|\)$/g, ''); cuenta.ocupacionCol = c + 1;
        } else if (t.length <= 40 && cuenta.notas.length < 3) {
          cuenta.notas.push(t);
        }
      });
    }
    // "reactivar" = cuenta dada de baja en el servicio: sus lugares no están disponibles
    // Inactivas: "reactivar", o "MIEMBRO EXTRA" sin número (con -0, -1, -2... está activa)
    // y "off" escrito en la fila de datos de la cuenta (cuentas apagadas)
    cuenta.inactiva = /reactivar/i.test(cuenta.ciclo) || norm_(cuenta.plan) === 'miembro extra' ||
      (infoRow != null && vals[infoRow].some(function (x) { return /^\s*off\s*$/i.test(String(x)); }));
    if (cuentaRow != null) {
      const av = vals[cuentaRow].filter(function (x) { return /aviso automatizado/i.test(x); });
      if (av.length) cuenta.aviso = String(av[0]).trim();
    }

    // Lugares
    for (let r = h + 1; r < vals.length; r++) {
      const row = vals[r];
      const a = norm_(row[0]);
      if (a === 'usuario' || esCuenta(row[0])) break;
      const usuario = limpio(row[col.usuario]);
      const fechaRaw = col.fecha != null ? String(row[col.fecha]).trim() : '';
      const conError = row.slice(1, anchoUtil).some(function (x) { return ES_ERROR.test(String(x).trim()); });
      if (!usuario && fechaRaw !== 'LIBRE' && !conError) break;
      // Fila pegada al encabezado de la cuenta siguiente sin datos del cliente: es la fila de datos de esa cuenta
      if (r + 1 < vals.length && norm_(vals[r + 1][0]) === 'usuario' && !conError && fechaRaw !== 'LIBRE' &&
          !limpio(row[col.id != null ? col.id : 2])) break;
      const g = function (k) { return col[k] != null ? limpio(row[col[k]]) : ''; };
      let msg = '';
      if (usuario) {
        if (top.data != null) msg = limpio(row[top.data]);
        if (!/:\*/.test(msg)) {
          msg = '';
          row.forEach(function (x) { const t = limpio(x); if (/:\*/.test(t) && t.length > msg.length && t.length > 40) msg = t; });
        }
      }
      const marcado = row.some(function (x, c) { return c > 0 && c < anchoUtil && /^\s*cerrar\s*$/i.test(String(x)); });
      // "comparte": usa el mismo perfil que otro cliente de la cuenta (ej. padre e hijo); no suma perfiles
      const comparte = row.some(function (x, c) { return c > 0 && c < anchoUtil && /^\s*comparte\s*$/i.test(String(x)); });
      if (!usuario) { cuenta.slots.push(marcado ? { row: r + 1, libre: true, cerrar: true } : { row: r + 1, libre: true }); continue; }
      const slot = {
        row: r + 1, usuario: usuario,
        huerfano: fechaRaw === 'LIBRE' || ES_ERROR.test(fechaRaw),
        id: g('id'), perfil: g('perfil'), numero: g('numero').replace(/\D/g, ''),
        fecha: fechaRaw === 'LIBRE' ? '' : limpio(fechaRaw), dias: g('dias'),
        monto: g('monto'), servicios: g('servicios'), pin: g('pin'),
        dispositivo: g('dispositivo'), correo: g('correo'), msg: msg, cerrar: marcado, comparte: comparte
      };
      Object.keys(slot).forEach(function (k) { if (slot[k] === '' || slot[k] === false) delete slot[k]; });
      cuenta.slots.push(slot);
    }
    if (!cuenta.nombre) {
      const enCab = head.filter(function (x) { return esCuenta(x); });
      if (enCab.length) cuenta.nombre = String(enCab[0]).trim();
    }
    cuenta._perfilCol = col.perfil != null ? col.perfil : null;
    cuentas.push(cuenta);
  }
  // Hay cuentas cuyo encabezado no dice "Perfil(es)" aunque la columna tenga la fórmula (pasa en HBO MAX).
  // Si en la hoja otras cuentas sí la tienen, se usa esa misma columna (solo valores chicos: 1 a 12).
  const usos = {};
  cuentas.forEach(function (c) { if (c._perfilCol != null) usos[c._perfilCol] = (usos[c._perfilCol] || 0) + 1; });
  let colP = null, maxP = 0;
  Object.keys(usos).forEach(function (k) { if (usos[k] > maxP) { maxP = usos[k]; colP = Number(k); } });
  cuentas.forEach(function (c) {
    if (colP != null && c._perfilCol == null) c.slots.forEach(function (s) {
      if (!s.usuario || s.perfil) return;
      const v = String((vals[s.row - 1] || [])[colP] == null ? '' : vals[s.row - 1][colP]).trim();
      if (/^\d{1,2}$/.test(v) && Number(v) >= 1 && Number(v) <= 12) s.perfil = v;
    });
    delete c._perfilCol;
  });
  return { hoja: hoja, cuentas: cuentas };
}

/* ------------------------------------------------------------------ */
/* Ayudantes                                                           */
/* ------------------------------------------------------------------ */

function ctx_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return { ss: ss, tz: ss.getSpreadsheetTimeZone() };
}

function conLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  PROPS_ALL_ = null;   // dentro del candado se leen las propiedades frescas (otro pudo guardar mientras esperábamos)
  try { return fn(); } finally { lock.releaseLock(); }
}

function norm_(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function hojaClientes_(ss) {
  const sh = ss.getSheetByName(CONFIG.HOJA_CLIENTES);
  if (!sh) throw new Error('No encuentro la hoja "' + CONFIG.HOJA_CLIENTES + '".');
  const ncol = sh.getLastColumn();
  // Velocidad: las columnas quedan guardadas 1 hora. Si agregás o borrás una columna (cambia la cantidad),
  // se vuelven a leer solas; si solo las cambiás de orden, ejecutá olvidarColumnas() desde el editor.
  const cache = CacheService.getScriptCache();
  try {
    const g = JSON.parse(cache.get(COLS_BASE_KEY_) || 'null');
    if (g && g.ncol === ncol && g.cols) return { sh: sh, cols: g.cols, ncol: ncol };
  } catch (e) { }
  const head = sh.getRange(1, 1, 1, ncol).getDisplayValues()[0];
  const cols = {};
  head.forEach(function (t, i) {
    const n = norm_(t);
    Object.keys(CAMPOS_CLIENTE).forEach(function (k) {
      if (cols[k] == null && CAMPOS_CLIENTE[k].indexOf(n) >= 0) cols[k] = i;
    });
  });
  ['nombre', 'fecha', 'meses'].forEach(function (k) {
    if (cols[k] == null) throw new Error('No encuentro la columna "' + k + '" en ' + CONFIG.HOJA_CLIENTES + '.');
  });
  try { cache.put(COLS_BASE_KEY_, JSON.stringify({ ncol: ncol, cols: cols }), 3600); } catch (e) { }
  return { sh: sh, cols: cols, ncol: ncol };
}
const COLS_BASE_KEY_ = 'colsBase_v1';
/** Ejecutala si cambiaste de orden las columnas de BasedeDatos (sin agregar ni borrar ninguna). */
function olvidarColumnas() {
  CacheService.getScriptCache().removeAll([COLS_BASE_KEY_, 'colsClientes_v1', 'leyenda_v1']);
  Logger.log('Listo: la app vuelve a leer los encabezados de BasedeDatos y la leyenda.');
}

function hojaServicio_(ss, hoja) {
  if (CONFIG.EXCLUIR.indexOf(hoja) >= 0) throw new Error('"' + hoja + '" no es una hoja de servicio.');
  const sh = ss.getSheetByName(hoja);
  if (!sh) throw new Error('No encuentro la hoja "' + hoja + '".');
  return sh;
}

function hojasServicio_(ss, forzar) {
  const cache = CacheService.getScriptCache();
  const KEY = 'hojasServicio_v1';
  const props = PropertiesService.getScriptProperties();
  if (!forzar) {
    const c = cache.get(KEY);
    if (c) return JSON.parse(c);
    // también queda guardada en el proyecto: las hojas casi nunca cambian, así no se revisan las 35 seguido
    const g = props.getProperty(KEY), cuando = Number(props.getProperty(KEY + '_ts') || 0);
    if (g && Date.now() - cuando < 24 * 3600000) { cache.put(KEY, g, 21600); return JSON.parse(g); }
  }
  const t0 = Date.now();
  const lista = hojasServicioRapido_(ss) || ss.getSheets().filter(function (sh) {
    if (sh.isSheetHidden() || CONFIG.EXCLUIR.indexOf(sh.getName()) >= 0) return false;
    const n = sh.getLastRow();
    if (n < 2) return false;
    const colA = sh.getRange(1, 1, n, 1).getDisplayValues();
    if (!colA.some(function (r) { return norm_(r[0]) === 'usuario'; })) return false;
    return !!sh.createTextFinder('DIAS RESTANTES').matchCase(false).findNext();
  }).map(function (sh) { return sh.getName(); });
  console.log('Lista de hojas de servicio revisada: %s ms', Date.now() - t0);
  cache.put(KEY, JSON.stringify(lista), 21600);
  try { props.setProperties({ hojasServicio_v1: JSON.stringify(lista), hojasServicio_v1_ts: String(Date.now()) }); } catch (e) { }
  return lista;
}

/**
 * Velocidad: la misma revisión que hojasServicio_, pero leyendo las columnas A:L de todas las hojas
 * en UN pedido (Google Sheets API) en vez de hoja por hoja (antes ~45 s, ahora ~1 s).
 * Si "DIAS RESTANTES" no aparece en A:L, esa hoja se revisa entera como antes, así el resultado es el mismo.
 * Devuelve null si no se puede (sin Sheets API): entonces se revisa hoja por hoja.
 */
function hojasServicioRapido_(ss) {
  if (!hayApiSheets_()) return null;
  try {
    const candidatas = ss.getSheets().filter(function (sh) {
      return !sh.isSheetHidden() && CONFIG.EXCLUIR.indexOf(sh.getName()) < 0;
    }).map(function (sh) { return sh.getName(); });
    if (!candidatas.length) return [];
    const r = Sheets.Spreadsheets.Values.batchGet(CONFIG.SPREADSHEET_ID, {
      ranges: candidatas.map(function (h) { return rangoHoja_(h) + '!A:L'; }), majorDimension: 'ROWS'
    });
    const vrs = r.valueRanges || [];
    if (vrs.length !== candidatas.length) return null;
    return candidatas.filter(function (hoja, i) {
      const filas = vrs[i].values || [];
      if (filas.length < 2) return false;
      if (!filas.some(function (f) { return norm_(f[0]) === 'usuario'; })) return false;
      const dias = filas.some(function (f) {
        return f.some(function (x) { return String(x == null ? '' : x).toUpperCase().indexOf('DIAS RESTANTES') >= 0; });
      });
      return dias || !!ss.getSheetByName(hoja).createTextFinder('DIAS RESTANTES').matchCase(false).findNext();
    });
  } catch (e) {
    console.warn('Revisar las hojas en un solo pedido falló, sigo hoja por hoja: ' + e);
    return null;
  }
}

function leyenda_(ss) {
  try {
    // la leyenda de emojis casi nunca cambia: queda guardada 6 horas (olvidarColumnas() la vuelve a leer)
    const cache = CacheService.getScriptCache();
    const g = cache.get('leyenda_v1');
    if (g) return JSON.parse(g);
    const sh = ss.getSheetByName(CONFIG.HOJA_REGISTRO);
    if (!sh) return [];
    const out = [];
    sh.getRange(CONFIG.RANGO_LEYENDA).getDisplayValues().forEach(function (r) {
      if (r[0] && r[1]) out.push([String(r[0]).trim(), String(r[1]).trim()]);
      if (r[3] && r[2]) out.push([String(r[3]).trim(), String(r[2]).trim()]);
    });
    try { cache.put('leyenda_v1', JSON.stringify(out), 21600); } catch (e) { }
    return out;
  } catch (e) { return []; }
}

function ultimaFila_(sh, colIdx) {
  const n = sh.getLastRow();
  if (n < 1) return 1;
  const v = sh.getRange(1, colIdx + 1, n, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]).trim() !== '') return i + 1;
  return 1;
}

function filaCliente_(h, row, nombre) {
  const n = String(nombre).trim();
  const r = Number(row);
  if (r >= 2 && r <= h.sh.getLastRow() && String(h.sh.getRange(r, h.cols.nombre + 1).getValue()).trim() === n) return r;
  const hits = filasConNombre_(h, function (v) { return v === n; });
  if (hits.length === 1) return hits[0];
  throw new Error('No encontré a "' + n + '" en la base. Tocá Actualizar.');
}
/**
 * Filas de BasedeDatos cuyo nombre cumple la condición. Lee la columna de nombres entera
 * (incluye las filas ocultas por un filtro: así funciona igual aunque alguien tenga la hoja filtrada).
 */
function filasConNombre_(h, cumple) {
  const n = h.sh.getLastRow();
  if (n < 2) return [];
  const vals = h.sh.getRange(2, h.cols.nombre + 1, n - 1, 1).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) if (cumple(String(vals[i][0]).trim())) out.push(i + 2);
  return out;
}
/**
 * Igual que filaCliente_, pero de paso devuelve la fila leída: { r, fila }.
 * Velocidad: en el caso normal (el cliente sigue en su fila) es UNA lectura en vez de tres.
 */
function filaYValores_(h, row, nombre) {
  const n = String(nombre).trim();
  const r0 = Number(row);
  if (r0 >= 2) {
    let v = null;
    try { v = h.sh.getRange(r0, 1, 1, h.ncol).getValues()[0]; } catch (e) { }   // fila fuera de la hoja: se busca abajo
    if (v && String(v[h.cols.nombre]).trim() === n) return { r: r0, fila: v };
  }
  const r = filaCliente_(h, row, nombre);   // se movió de fila: lo busca como antes
  return { r: r, fila: h.sh.getRange(r, 1, 1, h.ncol).getValues()[0] };
}

// VLOOKUP no distingue mayúsculas, así que la comparación tampoco.
function nombreExiste_(h, nombre, excluirFila) {
  const k = String(nombre).trim().toLowerCase();
  return filasConNombre_(h, function (v) { return v.toLowerCase() === k; })
    .some(function (r) { return r !== excluirFila; });
}

// Columnas que la app escribe: al crear una fila, el resto se copia con su fórmula.
function columnasManuales_(cols) {
  const m = {};
  editables_().forEach(function (k) { if (cols[k] != null) m[cols[k]] = true; });
  return m;
}

function textoLiteral_(s) {
  return /^[=+\-@]/.test(s) || /^\d+([.,]\d+)?$/.test(s) ? "'" + s : s;
}

function valorCampo_(k, v, tz) {
  const s = String(v == null ? '' : v).trim();
  switch (k) {
    case 'fecha':
      if (s.toUpperCase() === 'X') return 'X';
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return Utilities.parseDate(s, tz, 'yyyy-MM-dd');
      throw new Error('Fecha inválida: ' + s);
    case 'monto':
      if (!s) return '';
      if (s.toUpperCase() === 'X') return 'X';
      return Number(s.replace(/[^\d-]/g, '')) || 0;
    case 'meses':
      return Math.max(0, Math.round(Number(s) || 0));
    case 'perfil':
      return /^\d+$/.test(s) ? Number(s) : s;
    case 'telefono': {
      let d = s.replace(/\D/g, '');
      if (!d) return '';
      if (d.charAt(0) === '0') d = CONFIG.CODIGO_PAIS + d.slice(1);
      else if (d.length === 9 && d.charAt(0) === '9') d = CONFIG.CODIGO_PAIS + d;
      return Number(d);
    }
    case 'baja': case 'vencio':
      if (v instanceof Date) return v;
      if (!s) return '';
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return Utilities.parseDate(s, tz, 'yyyy-MM-dd');
      return '';
    case 'pin': case 'pins': case 'correo': case 'dispositivos':
      return s || CONFIG.VACIO;
    default:
      return s;
  }
}

function iguales_(a, b, tz) {
  if (a instanceof Date || b instanceof Date) {
    if (!(a instanceof Date && b instanceof Date)) return false;
    return Utilities.formatDate(a, tz, 'yyyy-MM-dd') === Utilities.formatDate(b, tz, 'yyyy-MM-dd');
  }
  return String(a).trim() === String(b).trim();
}

/**
 * Google ignora lo que un script escribe en una fila OCULTA POR UN FILTRO de la hoja (pasaba al reactivar
 * con BasedeDatos filtrada: decía que sí pero la Fecha quedaba en "x").
 * Si la fila está oculta por el filtro, lo saca un momento y devuelve una función que lo vuelve a poner
 * exactamente igual (mismo rango y mismas condiciones por columna). Si no hace falta, devuelve null.
 */
function filtroFuera_(sh, fila) {
  const f = sh.getFilter();
  if (!f) return null;
  let oculta = false;
  try { oculta = sh.isRowHiddenByFilter(fila); } catch (e) { }
  if (!oculta) return null;
  return quitarFiltro_(sh);
}
/** Saca el filtro de la hoja (si hay) y devuelve la función que lo vuelve a poner igual. */
function quitarFiltro_(sh) {
  const f = sh.getFilter();
  if (!f) return null;
  const rango = f.getRange();
  const c0 = rango.getColumn(), nc = rango.getNumColumns(), crit = [];
  for (let c = c0; c < c0 + nc; c++) { const x = f.getColumnFilterCriteria(c); if (x) crit.push([c, x.copy().build()]); }
  f.remove();
  console.log('Filtro de %s quitado un momento para escribir', sh.getName());
  return function () {
    try {
      SpreadsheetApp.flush();
      const nf = sh.getRange(rango.getA1Notation()).createFilter();
      crit.forEach(function (x) { nf.setColumnFilterCriteria(x[0], x[1]); });
    } catch (e) { console.warn('No se pudo volver a poner el filtro: ' + e); }
  };
}

/** Escribe solo los campos que cambiaron, y nunca en columnas calculadas. */
function escribirCampos_(h, r, datos, tz, filaLeida) {
  const restaurar = filtroFuera_(h.sh, r);   // fila oculta por un filtro: Google no la deja escribir
  try { return escribirCamposEn_(h, r, datos, tz, filaLeida); }
  finally { if (restaurar) restaurar(); }
}
function escribirCamposEn_(h, r, datos, tz, filaLeida) {
  const sh = h.sh;
  // si ya se leyó la fila en este mismo pedido (dentro del candado), no se vuelve a leer
  const actual = filaLeida || sh.getRange(r, 1, 1, h.ncol).getValues()[0];
  const cambios = [];
  const txt = function (v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'dd/MM/yyyy') : String(v == null ? '' : v).trim(); };
  editables_().forEach(function (k) {
    if (!(k in datos) || h.cols[k] == null) return;
    const c = h.cols[k];
    const nuevo = valorCampo_(k, datos[k], tz);
    if (iguales_(actual[c], nuevo, tz)) return;   // sin cambios: no se toca (ni su fórmula, si tuviera)
    const cell = sh.getRange(r, c + 1);
    if ((k === 'pin' || k === 'pins') && /^\d+$/.test(String(nuevo))) cell.setNumberFormat('@');
    cell.setValue(nuevo);
    cambios.push({ campo: k, antes: txt(actual[c]), despues: txt(nuevo), valor: nuevo });
  });
  return cambios;
}

function leerCliente_(h, r, tz) {
  SpreadsheetApp.flush();
  return enc_(h.sh.getRange(r, 1, 1, h.ncol).getValues()[0], r, h.cols, tz);
}

// Velocidad: formatDate es lento y muchas fechas se repiten (vencimientos del mismo día): cada fecha se formatea una sola vez
const FECHAS_ISO_ = {};
function fechaIso_(d, tz) {
  const k = tz + '|' + d.getTime();
  return FECHAS_ISO_[k] || (FECHAS_ISO_[k] = Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
}

function enc_(v, row, cols, tz) {
  const g = function (k) { return cols[k] == null ? '' : v[cols[k]]; };
  const f = function (x) {
    if (x instanceof Date) return fechaIso_(x, tz);
    return String(x == null ? '' : x).trim();
  };
  const tel = g('telefono');
  const monto = g('monto');
  const dias = g('dias');
  const meses = g('meses');
  return [
    row, f(g('nombre')), f(g('id')), f(g('perfil')), f(g('fecha')),
    typeof monto === 'number' ? monto : f(monto),
    typeof tel === 'number' ? String(Math.round(tel)) : f(tel),
    f(g('correo')), f(g('servicios')), f(g('dispositivos')), f(g('pin')), f(g('pins')),
    f(g('vencimiento')), typeof dias === 'number' ? Math.round(dias) : null,
    estadoKey_(g('estado')), typeof meses === 'number' ? meses : f(meses),
    f(g('baja')), f(g('vencio'))
  ];
}

function estadoKey_(s) {
  const t = norm_(s);
  if (t.indexOf('inactivo') === 0) return 'INACTIVO';
  if (t.indexOf('a vencer') === 0) return 'A VENCER';
  if (t.indexOf('vencido') === 0) return 'VENCIDO';
  if (t.indexOf('activo') === 0) return 'ACTIVO';
  return '';
}

function validarCeldaUsuario_(sh, row, col) {
  if (row < 2) throw new Error('Esa celda no parece un lugar de la columna Usuario.');
  const desde = Math.max(1, row - 15);
  const arriba = sh.getRange(desde, col, row - desde, 1).getDisplayValues();
  for (let i = arriba.length - 1; i >= 0; i--) {
    if (norm_(arriba[i][0]) === 'usuario') return;
  }
  throw new Error('Esa celda no parece un lugar de la columna Usuario.');
}

function renombrarEnServicios_(ss, viejo, nuevo) {
  let total = 0;
  const hojas = hojasServicio_(ss, false);
  // Velocidad: con Google Sheets API se reemplaza en la columna A de todas las hojas en UN pedido
  // (antes eran 35 búsquedas, una por hoja). Si falla, sigue hoja por hoja como antes.
  if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.batchUpdate) {
    try {
      const ids = {};
      ss.getSheets().forEach(function (sh) { ids[sh.getName()] = sh.getSheetId(); });
      const pedidos = hojas.filter(function (h) { return ids[h] != null; }).map(function (h) {
        return { findReplace: { find: viejo, replacement: nuevo, matchCase: true, matchEntireCell: true,
          range: { sheetId: ids[h], startColumnIndex: 0, endColumnIndex: 1 } } };
      });
      if (!pedidos.length) return 0;
      const r = Sheets.Spreadsheets.batchUpdate({ requests: pedidos }, CONFIG.SPREADSHEET_ID);
      (r.replies || []).forEach(function (x) { total += (x && x.findReplace && x.findReplace.occurrencesChanged) || 0; });
      return total;
    } catch (e) { console.warn('Renombrar en un solo pedido falló, sigo hoja por hoja: ' + e); total = 0; }
  }
  hojas.forEach(function (hoja) {
    const sh = ss.getSheetByName(hoja);
    if (!sh || sh.getLastRow() < 1) return;
    total += sh.getRange(1, 1, sh.getLastRow(), 1)
      .createTextFinder(viejo).matchEntireCell(true).matchCase(true).replaceAllWith(nuevo);
  });
  return total;
}

/* ------------------------------------------------------------------ */
/* Diagnóstico de fórmulas (solo lectura: no cambia nada de la planilla) */
/* ------------------------------------------------------------------ */

/**
 * Ejecutala UNA VEZ desde el editor. Recorre todas las hojas, cuenta las fórmulas y junta los tipos
 * que más pesan (las que buscan en columnas enteras, las que usan HOY/AHORA/INDIRECTO, etc.).
 * El resultado queda en una planilla NUEVA y chica en tu Drive: "Diagnóstico fórmulas AccountStore".
 * Tu planilla principal no se toca.
 */
function diagnosticarFormulas() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const VOLATIL = /\b(TODAY|NOW|HOY|AHORA|INDIRECT|INDIRECTO|OFFSET|DESREF|RAND|RANDBETWEEN|ALEATORIO)\s*\(/i;
  const BUSCA = /\b(VLOOKUP|BUSCARV|HLOOKUP|BUSCARH|XLOOKUP|BUSCARX|MATCH|COINCIDIR|INDEX|INDICE|COUNTIFS?|CONTAR\.SI|SUMIFS?|SUMAR\.SI|QUERY|FILTER|FILTRAR|UNIQUE|UNICOS|IMPORTRANGE|LOOKUP|BUSCAR)\s*\(/i;
  const ARRAY = /\b(ARRAYFORMULA|FORMULARRAY)\s*\(/i;
  // columnas enteras: A:A, 'Hoja'!B:F, C2:C (sin fila final), en notación R1C1: C[1]:C[1], C3:C5
  const COL_ENTERA = /(^|[^A-Z0-9])\$?[A-Z]{1,3}\$?\d*:\$?[A-Z]{1,3}(?![0-9A-Z])|C(\[-?\d+\])?\d*:C(\[-?\d+\])?\d*(?![0-9R])/;
  const resumen = [['Hoja', 'Oculta', 'Filas', 'Columnas', 'Celdas con fórmula', 'Usan HOY/AHORA/INDIRECTO', 'Buscan (BUSCARV, COINCIDIR, CONTAR.SI…)', 'Buscan en columnas enteras', 'ARRAYFORMULA', 'Segundos']];
  const patrones = [['Hoja', 'Veces', 'Ejemplo (celda)', 'Fórmula (R1C1: misma fórmula repetida en cada fila)']];
  ss.getSheets().forEach(function (sh) {
    const ti = Date.now();
    const nf = sh.getLastRow(), nc = sh.getLastColumn();
    if (nf < 1 || nc < 1) { resumen.push([sh.getName(), sh.isSheetHidden() ? 'sí' : '', nf, nc, 0, 0, 0, 0, 0, 0]); return; }
    const f = sh.getRange(1, 1, nf, nc).getFormulasR1C1();
    let total = 0, vol = 0, bus = 0, busEnt = 0, arr = 0;
    const cuenta = {};
    for (let r = 0; r < f.length; r++) for (let c = 0; c < f[r].length; c++) {
      const x = f[r][c];
      if (!x) continue;
      total++;
      if (VOLATIL.test(x)) vol++;
      const b = BUSCA.test(x);
      if (b) { bus++; if (COL_ENTERA.test(x)) busEnt++; }
      if (ARRAY.test(x)) arr++;
      const k = x.length > 400 ? x.slice(0, 400) + '…' : x;
      if (!cuenta[k]) cuenta[k] = { n: 0, celda: letraCol_(c + 1) + (r + 1) };
      cuenta[k].n++;
    }
    resumen.push([sh.getName(), sh.isSheetHidden() ? 'sí' : '', nf, nc, total, vol, bus, busEnt, arr, Math.round((Date.now() - ti) / 100) / 10]);
    Object.keys(cuenta).sort(function (a, b) { return cuenta[b].n - cuenta[a].n; }).slice(0, 12).forEach(function (k) {
      patrones.push([sh.getName(), cuenta[k].n, cuenta[k].celda, "'" + k]);
    });
  });
  const out = SpreadsheetApp.create('Diagnóstico fórmulas AccountStore');
  const h1 = out.getSheets()[0].setName('Resumen');
  h1.getRange(1, 1, resumen.length, resumen[0].length).setValues(resumen);
  const h2 = out.insertSheet('Fórmulas');
  h2.getRange(1, 1, patrones.length, patrones[0].length).setValues(patrones);
  Logger.log('Listo en %s segundos: %s', Math.round((Date.now() - t0) / 1000), out.getUrl());
}


/**
 * Ejecutala desde el editor para revisar a un cliente: escribí el nombre exacto abajo (entre comillas) y tocá Ejecutar.
 * Muestra en el registro todas las filas de BasedeDatos con ese nombre (fecha, meses, vencimiento) y
 * las últimas anotaciones de RegistroApp sobre él, y si la hoja tiene un filtro activo.
 */
function revisarCliente(nombre) {
  nombre = nombre || 'ESCRIBÍ ACÁ EL NOMBRE';
  const { ss, tz } = ctx_();
  const h = hojaClientes_(ss);
  const k = String(nombre).trim().toLowerCase();
  const filas = filasConNombre_(h, function (v) { return v.toLowerCase().indexOf(k) >= 0; });
  Logger.log('Filtro activo en BasedeDatos: %s', h.sh.getFilter() ? 'SÍ (' + h.sh.getFilter().getRange().getA1Notation() + ')' : 'no');
  Logger.log('Filas con "%s": %s', nombre, filas.length);
  filas.slice(0, 10).forEach(function (r) {
    const c = enc_(h.sh.getRange(r, 1, 1, h.ncol).getValues()[0], r, h.cols, tz);
    Logger.log('  fila %s: "%s" | fecha %s | meses %s | vence %s | días %s | oculta por filtro: %s', r, c[1], c[4], c[15], c[12], c[13],
      h.sh.isRowHiddenByFilter(r) ? 'sí' : 'no');
  });
  const log = ss.getSheetByName(CONFIG.HOJA_LOG);
  if (log && log.getLastRow() > 1) {
    const n = Math.min(400, log.getLastRow() - 1);
    log.getRange(log.getLastRow() - n + 1, 1, n, 5).getDisplayValues()
      .filter(function (x) { return String(x[3]).toLowerCase().indexOf(k) >= 0; }).slice(-5)
      .forEach(function (x) { Logger.log('  RegistroApp %s · %s · %s · %s', x[0], x[1], x[2], x[4]); });
  }
}

/**
 * Prueba de escritura en la Fecha de un cliente (para entender por qué no se guarda).
 * Escribí el número de fila abajo y ejecutala desde el editor. Muestra cómo está la celda
 * (fórmula, validación, protección, filtro) y prueba escribir la fecha de hoy y volver a leerla.
 * Si la fecha queda bien, el cliente queda activado desde hoy con 1 mes (igual que reactivarlo desde la app).
 */
function probarFechaFila(fila) {
  fila = Number(fila) || 1109;   // ← número de fila a probar (1109 = Rosalia Del Puerto📺)
  if (fila < 2) { Logger.log('Poné el número de fila en la primera línea de la función.'); return; }
  const { ss, tz } = ctx_();
  const h = hojaClientes_(ss);
  const sh = h.sh, col = h.cols.fecha + 1, cell = sh.getRange(fila, col);
  Logger.log('Columna Fecha = %s (encabezado "%s"), celda %s', col, sh.getRange(1, col).getDisplayValue(), cell.getA1Notation());
  Logger.log('Nombre: %s | valor: "%s" | fórmula: "%s" | formato: %s', sh.getRange(fila, h.cols.nombre + 1).getDisplayValue(),
    cell.getDisplayValue(), cell.getFormula(), cell.getNumberFormat());
  const dv = cell.getDataValidation();
  Logger.log('Validación: %s', dv ? dv.getCriteriaType() + ' ' + JSON.stringify(dv.getCriteriaValues()) + ' | permite inválidos: ' + dv.getAllowInvalid() : 'ninguna');
  Logger.log('Combinada: %s | oculta por filtro: %s', cell.isPartOfMerge(), sh.isRowHiddenByFilter(fila));
  const f = sh.getFilter();
  if (f) { const c = f.getColumnFilterCriteria(col); Logger.log('Filtro en la columna Fecha: %s', c ? (c.getCriteriaType() + ' ' + JSON.stringify(c.getCriteriaValues()) + ' ocultos: ' + JSON.stringify(c.getHiddenValues())) : 'ninguno'); }
  const prot = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(function (p) {
    const r = p.getRange(); return fila >= r.getRow() && fila <= r.getLastRow() && col >= r.getColumn() && col <= r.getLastColumn();
  });
  Logger.log('Rangos protegidos sobre la celda: %s', prot.length ? prot.map(function (p) { return p.getRange().getA1Notation() + ' (' + p.getDescription() + ')'; }).join(', ') : 'ninguno');
  const hoy = Utilities.parseDate(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), tz, 'yyyy-MM-dd');
  const restaurar = filtroFuera_(sh, fila);
  Logger.log(restaurar ? 'La fila estaba oculta por el filtro: lo saco un momento para escribir.' : 'No hace falta tocar el filtro.');
  cell.setValue(hoy);
  if (restaurar) restaurar();
  SpreadsheetApp.flush();
  Logger.log('Después de escribir la fecha de hoy, la celda dice: "%s"', cell.getDisplayValue());
}


/* ------------------------------------------------------------------ */
/* Bajas: fecha de baja (Baja) y vencimiento que tenía (Venció)        */
/* ------------------------------------------------------------------ */

/** Lo que se anota al dar de baja: hoy en Baja y el vencimiento que tenía en Venció. */
function datosDeBaja_(antes, h, tz) {
  if (h.cols.baja == null) return {};
  const d = { baja: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') };
  if (h.cols.vencio != null) d.vencio = /^\d{4}-\d{2}-\d{2}$/.test(String(antes[12])) ? antes[12] : '';
  return d;
}
/** Recargo por mora según los servicios (CONFIG.MORA_BAJA). 0 si no corresponde. */
function moraDe_(servicios) {
  const m = CONFIG.MORA_BAJA || {};
  let n = 0;
  String(servicios || '').split(',').forEach(function (sv) {
    Object.keys(m).forEach(function (k) { if (aplicaRegla_([k], sv.trim())) n = Math.max(n, Number(m[k]) || 0); });
  });
  return n;
}

/**
 * Ejecutala UNA VEZ desde el editor:
 *  1) agrega los encabezados "Baja" y "Venció" en BasedeDatos, después de MESES (si no están);
 *  2) completa las bajas que ya se hicieron desde la app (sale de RegistroApp);
 *  3) deja activo el aviso para las bajas que hagas a mano (X en Fecha).
 */
function prepararColumnasBaja() {
  const { ss, tz } = ctx_();
  let h = hojaClientes_(ss);
  const sh = h.sh;
  if (h.cols.baja == null) {
    const c = sh.getLastColumn() + 1;
    sh.getRange(1, h.cols.meses + 1).copyTo(sh.getRange(1, c, 1, 2), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sh.getRange(1, c, 1, 2).setValues([['Baja', 'Venció']]);
    sh.getRange(2, c, sh.getMaxRows() - 1, 2).setNumberFormat('dd/MM/yyyy').setHorizontalAlignment('center');
    CacheService.getScriptCache().remove(COLS_BASE_KEY_);
    h = hojaClientes_(ss);
    Logger.log('Columnas creadas: Baja (%s) y Venció (%s).', letraCol_(h.cols.baja + 1), letraCol_(h.cols.vencio + 1));
  } else Logger.log('Las columnas Baja y Venció ya existían.');

  // Bajas hechas desde la app: la última "Dio de baja" de cada cliente, con el vencimiento que tenía
  const log = ss.getSheetByName(CONFIG.HOJA_LOG), desdeLog = {};
  if (log && log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 5).getValues().forEach(function (x) {
    const k = String(x[3]).trim().toLowerCase();
    if (!k || !(x[0] instanceof Date)) return;
    if (x[2] === 'Dio de baja') {
      const m = String(x[4]).match(/Vencimiento: (\d{2})\/(\d{2})\/(\d{4})/);
      desdeLog[k] = { baja: x[0], vencio: m ? m[3] + '-' + m[2] + '-' + m[1] : '' };
    } else if (x[2] === 'Reactivó') delete desdeLog[k];
  });
  const last = ultimaFila_(sh, h.cols.nombre);
  if (last > 1) {
    const nom = sh.getRange(2, h.cols.nombre + 1, last - 1, 1).getValues();
    const fec = sh.getRange(2, h.cols.fecha + 1, last - 1, 1).getValues();
    const rq = sh.getRange(2, h.cols.baja + 1, last - 1, 2);
    const qr = rq.getValues();
    let n = 0;
    nom.forEach(function (x, i) {
      const k = String(x[0]).trim().toLowerCase(), info = desdeLog[k];
      if (!info || String(fec[i][0]).trim().toUpperCase() !== 'X' || qr[i][0] !== '') return;
      qr[i][0] = info.baja;
      qr[i][1] = info.vencio ? Utilities.parseDate(info.vencio, tz, 'yyyy-MM-dd') : '';
      n++;
    });
    if (n) {
      const restaurar = quitarFiltro_(sh);
      try { rq.setValues(qr); } finally { if (restaurar) restaurar(); }
    }
    Logger.log('Bajas completadas desde RegistroApp: %s. Las que se hicieron a mano antes quedan sin fecha.', n);
  }
  activarAnotarBajas_();
}

/** Deja activo el aviso de la planilla: cuando alguien pone X en Fecha a mano, se anotan Baja y Venció. */
function activarAnotarBajas_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'alEditarPlanilla') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('alEditarPlanilla').forSpreadsheet(CONFIG.SPREADSHEET_ID).onEdit().create();
  Logger.log('Listo: cuando pongas X en Fecha a mano, se anotan solas la fecha de baja y el vencimiento que tenía.');
}

/**
 * Se ejecuta sola cada vez que alguien edita la planilla a mano (no cuando escribe la app).
 * Solo actúa si se cambió UNA celda de Fecha en BasedeDatos:
 *  - de una fecha a X → Baja = hoy, Venció = fecha anterior + MESES (y en Tigo Sports no toca nada más: el -5000 lo ponés vos)
 *  - de X a una fecha (reactivado a mano) → limpia Baja y Venció
 */
function alEditarPlanilla(e) {
  try {
    const rg = e && e.range;
    if (!rg || rg.getNumRows() !== 1 || rg.getNumColumns() !== 1 || rg.getRow() < 2) return;
    const sh = rg.getSheet();
    if (sh.getName() !== CONFIG.HOJA_CLIENTES) return;
    const ss = sh.getParent(), h = hojaClientes_(ss);
    if (h.cols.baja == null || rg.getColumn() !== h.cols.fecha + 1) return;
    const tz = ss.getSpreadsheetTimeZone(), r = rg.getRow();
    const nuevo = String(e.value == null ? rg.getDisplayValue() : e.value).trim();
    const viejo = String(e.oldValue == null ? '' : e.oldValue).trim();
    const celdas = sh.getRange(r, h.cols.baja + 1, 1, 2);
    if (nuevo.toUpperCase() === 'X' && viejo.toUpperCase() !== 'X') {
      const f = fechaDeEvento_(viejo);
      const meses = Number(sh.getRange(r, h.cols.meses + 1).getValue());
      const venc = f && meses >= 0 ? edateIso_(f, meses) : '';
      celdas.setValues([[new Date(), venc ? Utilities.parseDate(venc, tz, 'yyyy-MM-dd') : '']]);
      anotarFechaBaja_(sh.getRange(r, h.cols.nombre + 1).getDisplayValue(), 'Dio de baja');
    } else if (viejo.toUpperCase() === 'X' && nuevo && nuevo.toUpperCase() !== 'X') {
      celdas.clearContent();
    }
  } catch (err) { console.warn('alEditarPlanilla: ' + err); }
}
/** La fecha anterior que trae el aviso de edición: número de serie (45930) o texto (19/08/2026). Devuelve yyyy-MM-dd o ''. */
function fechaDeEvento_(v) {
  if (/^\d+(\.\d+)?$/.test(v)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(v)) * 864e5);
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }
  let m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) { const y = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]); return y + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0'); }
  m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[1] + '-' + m[2] + '-' + m[3] : '';
}
