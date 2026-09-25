/**
 * VetGestión · NeoZoo — conexión entre el programa y esta planilla de Google Sheets.
 *
 * Cómo instalarlo (una sola vez):
 *   1. En la planilla: Extensiones > Apps Script. Borrá lo que haya y pegá este código. Guardá.
 *   2. Implementar > Nueva implementación > tipo "Aplicación web".
 *        Ejecutar como: Yo      ·      Quién tiene acceso: Cualquier persona
 *        (la última opción de la lista, la que NO dice "con una cuenta de Google")
 *   3. Autorizá los permisos (si aparece "Google no verificó esta app": Configuración avanzada > Ir a...).
 *   4. Copiá la URL que termina en /exec y pegala en VetGestión > Administración > Base de datos.
 */
const ID_PLANILLA = '';  // vacío = la planilla donde está pegado el script; o el ID de una planilla
const HOJA_PROD = 'Productos';
const HOJA_MOV = 'Movimientos';
const CARPETA_IMAGENES = 'VetGestión - imágenes';

function doGet() {
  return salida({ok: true, mensaje: 'VetGestión: el script está funcionando.', nombre: libro().getName()});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);  // una operación a la vez, aunque escriban varias computadoras
  try {
    const pedido = JSON.parse(e.postData.contents);
    const accion = ACCIONES[pedido.accion];
    if (!accion) throw new Error('Acción desconocida: ' + pedido.accion);
    return salida(Object.assign({ok: true}, accion(pedido) || {}));
  } catch (err) {
    return salida({ok: false, error: String((err && err.message) || err), usuario: !!(err && err.usuario)});
  } finally {
    lock.releaseLock();
  }
}

const ACCIONES = {
  ping: function () {
    const h = libro().getSheetByName(HOJA_PROD);
    return {nombre: libro().getName(), url: libro().getUrl(), productos: h ? Math.max(h.getLastRow() - 1, 0) : 0,
            version: VERSION};
  },

  /* Crea las hojas con encabezados y formato. Solo carga filas si la hoja Productos está vacía. */
  inicializar: function (p) {
    const hp = prepararHoja(HOJA_PROD, p.prod);
    const hm = prepararHoja(HOJA_MOV, p.mov);
    let subidos = false;
    if (hp.getLastRow() <= 1 && p.prod.filas.length) {
      escribir(hp, 2, p.prod.filas);
      if (hm.getLastRow() <= 1 && p.mov.filas.length) escribir(hm, 2, p.mov.filas);
      subidos = true;
    }
    // la hoja vacía que trae toda planilla nueva ("Hoja 1") ya no hace falta
    libro().getSheets().forEach(function (h) {
      if (h.getName() !== HOJA_PROD && h.getName() !== HOJA_MOV && h.getLastRow() === 0 && libro().getSheets().length > 2) {
        libro().deleteSheet(h);
      }
    });
    return {subidos: subidos};
  },

  leer: function () {
    return {productos: filas(HOJA_PROD), movimientos: filas(HOJA_MOV), version: VERSION};
  },

  /* parciales: cambios puntuales {codigo, stock (a sumar/restar), precio (nuevo)} que se aplican sobre
       lo que hay AHORA en la planilla, así dos computadoras vendiendo a la vez no se pisan.
     upserts: filas completas de productos (se reemplaza la fila con el mismo código o se agrega).
     borrar: códigos a eliminar. movs: filas nuevas del historial.
     devolver: si es true, responde con los datos actualizados (ahorra una lectura). */
  aplicar: function (p) {
    const h = libro().getSheetByName(HOJA_PROD);
    const n = h.getLastRow() - 1;
    const codigos = n > 0 ? h.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
    const colStock = p.colStock || 9, colPrecio = p.colPrecio || 8;
    // primero se valida todo, para no dejar cambios a medias si algo falla
    const parciales = (p.parciales || []).map(function (c) {
      const i = codigos.indexOf(String(c.codigo).trim());
      if (i < 0) throw aviso('El producto ' + c.codigo + ' ya no existe.');
      const r = {fila: i + 2, precio: c.precio};
      if (c.stock) {
        const actual = Number(h.getRange(i + 2, colStock).getValue()) || 0;
        r.stock = actual + Number(c.stock);
        if (r.stock < 0) throw aviso('No alcanza el stock: quedan ' + actual + ' unidades.');
      }
      return r;
    });
    parciales.forEach(function (r) {
      if (r.stock !== undefined) h.getRange(r.fila, colStock).setValue(r.stock);
      if (r.precio !== undefined && r.precio !== null) h.getRange(r.fila, colPrecio).setValue(r.precio);
    });
    (p.upserts || []).forEach(function (fila) {
      const i = codigos.indexOf(String(fila[0]).trim());
      if (i >= 0) {
        h.getRange(i + 2, 1, 1, fila.length).setValues([fila]);
      } else {
        h.appendRow(fila);
        codigos.push(String(fila[0]).trim());
      }
    });
    (p.borrar || [])
      .map(function (c) { return codigos.indexOf(String(c).trim()); })
      .filter(function (i) { return i >= 0; })
      .sort(function (a, b) { return b - a; })
      .forEach(function (i) { h.deleteRow(i + 2); });
    const hm = libro().getSheetByName(HOJA_MOV);
    (p.movs || []).forEach(function (fila) { hm.appendRow(fila); });
    return p.devolver ? ACCIONES.leer() : {};
  },

  subir_imagen: function (p) {
    const it = DriveApp.getFoldersByName(CARPETA_IMAGENES);
    const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_IMAGENES);
    const blob = Utilities.newBlob(Utilities.base64Decode(p.datos), p.mime, p.nombre);
    return {id: carpeta.createFile(blob).getId()};
  },

  imagen: function (p) {
    const blob = DriveApp.getFileById(p.id).getBlob();
    return {mime: blob.getContentType(), datos: Utilities.base64Encode(blob.getBytes())};
  }
};

/* ---------- auxiliares ---------- */
const VERSION = 2;

/* Error para mostrarle tal cual a quien usa el programa (no es una falla del script). */
function aviso(mensaje) {
  const e = new Error(mensaje);
  e.usuario = true;
  return e;
}

function libro() {
  return ID_PLANILLA ? SpreadsheetApp.openById(ID_PLANILLA) : SpreadsheetApp.getActiveSpreadsheet();
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function prepararHoja(nombre, def) {
  let h = libro().getSheetByName(nombre);
  if (!h) h = libro().insertSheet(nombre);
  const cols = def.encabezados.length;
  if (h.getMaxColumns() < cols) h.insertColumnsAfter(h.getMaxColumns(), cols - h.getMaxColumns());
  h.getRange(1, 1, 1, cols).setValues([def.encabezados])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#123a30');
  h.setFrozenRows(1);
  // texto plano para que "08/2026" o "005-0142" no se conviertan en fechas; precios y cantidades como número
  def.formatos.forEach(function (fmt, i) {
    h.getRange(1, i + 1, h.getMaxRows(), 1).setNumberFormat(fmt);
    if (def.anchos) h.setColumnWidth(i + 1, def.anchos[i]);
  });
  return h;
}

function escribir(h, desde, datos) {
  const faltan = desde + datos.length - 1 - h.getMaxRows();
  if (faltan > 0) h.insertRowsAfter(h.getMaxRows(), faltan);
  h.getRange(desde, 1, datos.length, datos[0].length).setValues(datos);
}

function filas(nombre) {
  const h = libro().getSheetByName(nombre);
  if (!h || h.getLastRow() < 2) return [];
  const zona = Session.getScriptTimeZone();
  return h.getRange(2, 1, h.getLastRow() - 1, h.getLastColumn()).getValues().map(function (fila) {
    return fila.map(function (v) {
      return v instanceof Date ? Utilities.formatDate(v, zona, 'yyyy-MM-dd HH:mm:ss') : v;
    });
  });
}
