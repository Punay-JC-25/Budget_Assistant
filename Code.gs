/**
 * Cuentas con fecha de corte y de pago - backend en Google Apps Script
 *
 * Usa dos hojas dentro de este mismo archivo de Google Sheets:
 *   - "Cuentas":     cada tarjeta o pago, con sus días de corte y de pago.
 *   - "Movimientos": cada cantidad registrada (fecha, hora, cantidad, usuario, cuenta).
 *
 * Días de corte: el último día del periodo de compras. Puede haber varios por mes
 * (por ejemplo "4,19" para una cuenta quincenal).
 * Días de pago: opcionales; se muestra el siguiente día de pago después del corte.
 */

const HOJA_MOV = "Movimientos";
const HOJA_CTA = "Cuentas";
const ZONA = "America/Guatemala";

// ---------- Entradas web ----------

// Devuelve las cuentas con su periodo actual y el total acumulado
function doGet(e) {
  try {
    return json_(estado_());
  } catch (err) {
    return json_({ ok: false, error: msg_(err) });
  }
}

// Recibe: registros (también los guardados sin internet), crear/editar cuenta u ocultar cuenta
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    let cuentaId = null;
    if (d.accion === "guardarCuenta") cuentaId = guardarCuenta_(d.cuenta);
    else if (d.accion === "ocultarCuenta") ocultarCuenta_(d.id);
    else registrar_(d.entries || []);
    const res = estado_();
    if (cuentaId) res.cuentaId = cuentaId;
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: msg_(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Estado y totales ----------

function estado_() {
  const hoy = Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd");
  const cuentas = leerCuentas_().filter(function (c) { return c.activa; });
  const mov = leerMovimientos_();
  const salida = cuentas.map(function (c) {
    const p = periodoDe_(hoy, c.cortes, c.pagos);
    let total = 0;
    mov.forEach(function (m) {
      if (m.cuenta === c.id && m.dia >= p.inicio && m.dia <= p.fin) total += m.monto;
    });
    return {
      id: c.id, nombre: c.nombre, cortes: c.cortes, pagos: c.pagos,
      inicio: p.inicio, fin: p.fin, pagoFecha: p.pago,
      total: Math.round(total * 100) / 100
    };
  });
  return { ok: true, hoy: hoy, cuentas: salida };
}

// ---------- Movimientos ----------

function registrar_(entries) {
  const cuentas = {};
  leerCuentas_().forEach(function (c) { cuentas[c.id] = c; });
  const sh = hojaMov_();
  const yaGuardados = idsExistentes_(sh);
  const filas = [];

  entries.forEach(function (en) {
    if (!en || !en.id || yaGuardados[String(en.id)]) return;     // evita duplicados si se reenvía
    const c = cuentas[String(en.cuentaId)];
    if (!c) return;
    const monto = Number(en.monto);
    if (!isFinite(monto) || monto <= 0) return;
    const fecha = new Date(en.ts);
    if (isNaN(fecha.getTime())) return;
    filas.push([
      String(en.id),
      fecha,
      monto,
      String(en.usuario || "").slice(0, 60),
      c.id,
      c.nombre,
      Utilities.formatDate(fecha, ZONA, "yyyy-MM-dd")
    ]);
    yaGuardados[String(en.id)] = true;
  });

  if (filas.length) {
    sh.getRange(sh.getLastRow() + 1, 1, filas.length, 7).setValues(filas);
  }
}

function leerMovimientos_() {
  const sh = hojaMov_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 7).getValues().map(function (f) {
    return { monto: Number(f[2]) || 0, cuenta: String(f[4]), dia: String(f[6]) };
  });
}

function idsExistentes_(sh) {
  const mapa = {};
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    sh.getRange(2, 1, n, 1).getValues().forEach(function (f) { mapa[String(f[0])] = true; });
  }
  return mapa;
}

// ---------- Cuentas ----------

function leerCuentas_() {
  const sh = hojaCta_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 5).getValues().map(function (f) {
    return {
      id: String(f[0]),
      nombre: String(f[1]),
      cortes: limpiarDias_(f[2]),
      pagos: limpiarDias_(f[3]),
      activa: String(f[4]).toUpperCase() !== "NO"
    };
  }).filter(function (c) { return c.id && c.cortes.length; });
}

function guardarCuenta_(c) {
  if (!c) throw new Error("Falta la cuenta");
  const nombre = String(c.nombre || "").trim().slice(0, 40);
  if (!nombre) throw new Error("Escribe un nombre para la cuenta");
  const cortes = limpiarDias_(c.cortes);
  if (!cortes.length) throw new Error("Indica al menos un día de corte (de 1 a 31)");
  const pagos = limpiarDias_(c.pagos);
  const sh = hojaCta_();

  if (c.id) {
    const fila = buscarFilaCuenta_(sh, c.id);
    if (fila) {
      sh.getRange(fila, 2, 1, 4).setValues([[nombre, cortes.join(","), pagos.join(","), "SI"]]);
      return String(c.id);
    }
  }
  const id = Utilities.getUuid().slice(0, 8);
  sh.appendRow([id, nombre, cortes.join(","), pagos.join(","), "SI"]);
  return id;
}

function ocultarCuenta_(id) {
  const sh = hojaCta_();
  const fila = buscarFilaCuenta_(sh, id);
  if (!fila) throw new Error("No se encontró la cuenta");
  sh.getRange(fila, 5).setValue("NO");
}

function buscarFilaCuenta_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

// Acepta una lista [4,19] o un texto "4, 19"; devuelve días únicos de 1 a 31, ordenados
function limpiarDias_(v) {
  const partes = Array.isArray(v) ? v : String(v == null ? "" : v).split(/[^0-9]+/);
  const vistos = {};
  const dias = [];
  partes.forEach(function (x) {
    const n = Number(x);
    if (x !== "" && x !== null && isFinite(n) && n >= 1 && n <= 31 && Math.floor(n) === n && !vistos[n]) {
      vistos[n] = true;
      dias.push(n);
    }
  });
  dias.sort(function (a, b) { return a - b; });
  return dias.slice(0, 6);
}

// ---------- Hojas ----------

function hojaMov_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA_MOV);
  if (!sh) {
    sh = ss.insertSheet(HOJA_MOV);
    sh.getRange("A:A").setNumberFormat("@");
    sh.getRange("E:E").setNumberFormat("@");
    sh.getRange("G:G").setNumberFormat("@");
    sh.getRange("B:B").setNumberFormat("dd/mm/yyyy hh:mm:ss");
    sh.getRange(1, 1, 1, 7)
      .setValues([["ID", "Fecha y hora", "Cantidad", "Usuario", "ID cuenta", "Cuenta", "Día"]])
      .setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function hojaCta_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA_CTA);
  if (!sh) {
    sh = ss.insertSheet(HOJA_CTA);
    sh.getRange("A:E").setNumberFormat("@");   // texto, para que "4,19" no se convierta en número
    sh.getRange(1, 1, 1, 5)
      .setValues([["ID", "Nombre", "Días de corte", "Días de pago", "Activa"]])
      .setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

// ---------- Periodos (misma lógica que index.html) ----------

function pad_(n) { return (n < 10 ? "0" : "") + n; }

function diasEnMes_(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }   // m de 1 a 12

function sumaMes_(y, m, k) {
  const t = y * 12 + (m - 1) + k;
  return [Math.floor(t / 12), (t % 12) + 1];
}

function aIso_(dt) {
  return dt.getUTCFullYear() + "-" + pad_(dt.getUTCMonth() + 1) + "-" + pad_(dt.getUTCDate());
}

// Fechas AAAA-MM-DD de los días indicados, desde el mes (m+desde) hasta el mes (m+hasta)
function fechasCandidatas_(dias, y, m, desde, hasta) {
  const salida = [];
  for (let k = desde; k <= hasta; k++) {
    const ym = sumaMes_(y, m, k);
    dias.forEach(function (d) {
      salida.push(ym[0] + "-" + pad_(ym[1]) + "-" + pad_(Math.min(d, diasEnMes_(ym[0], ym[1]))));
    });
  }
  const unicas = [];
  salida.sort().forEach(function (f) { if (unicas[unicas.length - 1] !== f) unicas.push(f); });
  return unicas;
}

// hoy = "AAAA-MM-DD". Devuelve {inicio, fin, pago}. El periodo termina en el primer corte >= hoy.
function periodoDe_(hoy, cortes, pagos) {
  const p = hoy.split("-").map(Number);
  const cortesF = fechasCandidatas_(cortes, p[0], p[1], -1, 1);
  const fin = cortesF.filter(function (f) { return f >= hoy; })[0];
  const previos = cortesF.filter(function (f) { return f < fin; });
  const prev = previos[previos.length - 1].split("-").map(Number);
  const inicio = aIso_(new Date(Date.UTC(prev[0], prev[1] - 1, prev[2] + 1)));
  let pago = null;
  if (pagos && pagos.length) {
    const f = fin.split("-").map(Number);
    const pagosF = fechasCandidatas_(pagos, f[0], f[1], 0, 2);
    pago = pagosF.filter(function (x) { return x > fin; })[0] || null;
  }
  return { inicio: inicio, fin: fin, pago: pago };
}

// ---------- Respuestas ----------

function msg_(err) { return String(err && err.message ? err.message : err); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
