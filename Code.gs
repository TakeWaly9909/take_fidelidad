/**
 * TAKE AWAY — API de fidelización sobre Google Sheets
 * ---------------------------------------------------
 * 1. Crea una Google Sheet con una hoja llamada "Clientes" y estos encabezados
 *    en la fila 1: telefono | nombre | puntos | sellos | fecha
 * 2. Extensiones > Apps Script, pega este archivo reemplazando el contenido.
 * 3. Corre UNA VEZ la función configurarPin() (▶ arriba, elige esa función)
 *    para guardar el PIN del admin. Cámbialo antes de correrlo.
 * 4. Implementar > Nueva implementación > Tipo: Aplicación web
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquiera
 *    Copia la URL que te da, la vas a usar en los 3 HTML.
 */

function configurarPin() {
  // Cambia '1234' por el PIN real antes de correr esta función una sola vez.
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', '1234');
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
}

function buscarFila(sheet, telefono) {
  var datos = sheet.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(telefono)) return i + 1;
  }
  return null;
}

function buscarCliente(sheet, telefono) {
  var fila = buscarFila(sheet, telefono);
  if (!fila) return null;
  var datos = sheet.getRange(fila, 1, 1, 4).getValues()[0];
  return { telefono: String(datos[0]), nombre: datos[1], puntos: datos[2], sellos: datos[3] };
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var accion = e.parameter.accion;
    var sheet = getSheet();

    if (accion === 'consultar') {
      var cliente = buscarCliente(sheet, e.parameter.telefono);
      return salida(cliente || { error: 'no_encontrado' });
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
      if (buscarCliente(sheet, data.telefono)) {
        return salida({ error: 'ya_existe' });
      }
      sheet.appendRow([data.telefono, data.nombre, 0, 0, new Date()]);
      return salida({ ok: true });
    }

    if (data.accion === 'sumar' || data.accion === 'canjear') {
      var pinGuardado = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
      if (data.pin !== pinGuardado) {
        return salida({ error: 'pin_incorrecto' });
      }
      var fila = buscarFila(sheet, data.telefono);
      if (!fila) return salida({ error: 'no_encontrado' });

      var puntosActuales = sheet.getRange(fila, 3).getValue();
      var sellosActuales = sheet.getRange(fila, 4).getValue();
      var signo = data.accion === 'sumar' ? 1 : -1;

      var nuevosPuntos = Math.max(0, puntosActuales + signo * (data.puntos || 0));
      var nuevosSellos = Math.max(0, Math.min(10, sellosActuales + signo * (data.sellos || 0)));

      sheet.getRange(fila, 3).setValue(nuevosPuntos);
      sheet.getRange(fila, 4).setValue(nuevosSellos);

      return salida({ ok: true, puntos: nuevosPuntos, sellos: nuevosSellos });
    }

    return salida({ error: 'accion_invalida' });
  } catch (err) {
    return salida({ error: err.message });
  }
}
