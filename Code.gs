

var SESION_TTL_SEG = 1800;        // duración de la sesión del cajero (30 min)
var MAX_INTENTOS_LOGIN = 5;       // intentos de PIN antes de bloquear
var VENTANA_INTENTOS_SEG = 300;   // ventana del bloqueo (5 min)
var CACHE_CONSULTA_SEG = 20;      // cuánto se cachea una consulta de cliente
var CACHE_LISTA_SEG = 15;         // cuánto se cachea la tabla completa de clientes

function configurarPin() {
 
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN_HASH', hashHex_('9909'));
}

function hashHex_(texto) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function pinEsDefecto_(hash) {
  return hash === hashHex_('1234');
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
}

// Solo lee la columna de teléfono (no toda la hoja) para ubicar la fila.
function buscarFila(sheet, telefono) {
  var ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) return null;
  var telefonos = sheet.getRange(2, 1, ultimaFila - 1, 1).getValues();
  for (var i = 0; i < telefonos.length; i++) {
    if (String(telefonos[i][0]) === String(telefono)) return i + 2;
  }
  return null;
}

function buscarCliente(sheet, telefono) {
  var fila = buscarFila(sheet, telefono);
  if (!fila) return null;
  var datos = sheet.getRange(fila, 1, 1, 4).getValues()[0];
  return { telefono: String(datos[0]), nombre: datos[1], puntos: datos[2], sellos: datos[3] };
}

// Trae toda la hoja en una sola lectura (para el panel de admin).
function listarClientes(sheet) {
  var ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) return [];
  var datos = sheet.getRange(2, 1, ultimaFila - 1, 5).getValues();
  var lista = [];
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    lista.push({
      telefono: String(datos[i][0]),
      nombre: datos[i][1],
      puntos: datos[i][2],
      sellos: datos[i][3],
      fecha: datos[i][4] instanceof Date ? datos[i][4].toISOString() : String(datos[i][4])
    });
  }
  return lista;
}

function cache_() {
  return CacheService.getScriptCache();
}

function invalidarCacheCliente_(telefono) {
  var cache = cache_();
  cache.remove('cliente_' + telefono);
  cache.remove('lista_clientes');
}

function loginBloqueado_() {
  return Number(cache_().get('login_intentos') || 0) >= MAX_INTENTOS_LOGIN;
}

function registrarIntentoFallido_() {
  var cache = cache_();
  var intentos = Number(cache.get('login_intentos') || 0) + 1;
  cache.put('login_intentos', String(intentos), VENTANA_INTENTOS_SEG);
}

function limpiarIntentos_() {
  cache_().remove('login_intentos');
}

function crearSesion_() {
  var token = Utilities.getUuid();
  cache_().put('sesion_' + token, 'valida', SESION_TTL_SEG);
  return token;
}

function sesionValida_(token) {
  return !!token && cache_().get('sesion_' + token) === 'valida';
}

function invalidarSesion_(token) {
  if (token) cache_().remove('sesion_' + token);
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var accion = e.parameter.accion;

    if (accion === 'consultar') {
      var telefono = e.parameter.telefono;
      var cache = cache_();
      var clave = 'cliente_' + telefono;
      var enCache = cache.get(clave);
      if (enCache) return salida(JSON.parse(enCache));

      var cliente = buscarCliente(getSheet(), telefono);
      if (!cliente) return salida({ error: 'no_encontrado' });

      cache.put(clave, JSON.stringify(cliente), CACHE_CONSULTA_SEG);
      return salida(cliente);
    }
    return salida({ error: 'accion_invalida' });
  } catch (err) {
    return salida({ error: err.message });
  }
}

function doPost(e) {
  try {
    // Se envía como text/plain para evitar el preflight CORS de fetch.
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet();

    if (data.accion === 'registro') {
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        if (buscarCliente(sheet, data.telefono)) {
          return salida({ error: 'ya_existe' });
        }
        sheet.appendRow([data.telefono, data.nombre, 0, 0, new Date()]);
      } finally {
        lock.releaseLock();
      }
      invalidarCacheCliente_(data.telefono);
      return salida({ ok: true });
    }

    if (data.accion === 'login') {
      if (loginBloqueado_()) {
        return salida({ error: 'demasiados_intentos' });
      }
      var pinGuardadoHash = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN_HASH');
      if (data.pinHash !== pinGuardadoHash) {
        registrarIntentoFallido_();
        return salida({ error: 'pin_incorrecto' });
      }
      limpiarIntentos_();
      var respuesta = { ok: true, token: crearSesion_() };
      if (pinEsDefecto_(pinGuardadoHash)) respuesta.warning = 'pin_default';
      return salida(respuesta);
    }

    if (data.accion === 'logout') {
      invalidarSesion_(data.token);
      return salida({ ok: true });
    }

    if (data.accion === 'listar') {
      if (!sesionValida_(data.token)) {
        return salida({ error: 'sesion_invalida' });
      }
      var cache = cache_();
      var enCache = cache.get('lista_clientes');
      if (enCache) return salida(JSON.parse(enCache));

      var clientes = listarClientes(sheet);
      try {
        cache.put('lista_clientes', JSON.stringify(clientes), CACHE_LISTA_SEG);
      } catch (errCache) {
        // La lista es demasiado grande para el cache; no pasa nada, se sirve sin cachear.
      }
      return salida(clientes);
    }

    if (data.accion === 'sumar' || data.accion === 'canjear') {
      if (!sesionValida_(data.token)) {
        return salida({ error: 'sesion_invalida' });
      }
      var fila = buscarFila(sheet, data.telefono);
      if (!fila) return salida({ error: 'no_encontrado' });

      var lockEdicion = LockService.getScriptLock();
      lockEdicion.waitLock(10000);
      var nuevosPuntos, nuevosSellos;
      try {
        var actuales = sheet.getRange(fila, 3, 1, 2).getValues()[0];
        var signo = data.accion === 'sumar' ? 1 : -1;
        nuevosPuntos = Math.max(0, actuales[0] + signo * (data.puntos || 0));
        nuevosSellos = Math.max(0, Math.min(10, actuales[1] + signo * (data.sellos || 0)));
        sheet.getRange(fila, 3, 1, 2).setValues([[nuevosPuntos, nuevosSellos]]);
      } finally {
        lockEdicion.releaseLock();
      }

      invalidarCacheCliente_(data.telefono);
      return salida({ ok: true, puntos: nuevosPuntos, sellos: nuevosSellos });
    }

    return salida({ error: 'accion_invalida' });
  } catch (err) {
    return salida({ error: err.message });
  }
}
