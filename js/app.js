/* =====================================================================
   Automatización Contable: Libro Diario → Libro Mayor → Balance 8 Columnas
   Todo el procesamiento ocurre en el navegador (sin servidor).
   ===================================================================== */
"use strict";

/* ---------------- Estado global ---------------- */
let entries = [];            // [{num, fecha, glosa, lineas:[{cuenta, debe, haber}]}]
let alertaFuente = null;     // mensaje sobre el origen de los datos (se muestra junto a las validaciones)
const clasifOverrides = {};  // nombreNormalizado -> 'A' | 'P' | 'PE' | 'GA'

const TIPOS = { A: "Activo", P: "Pasivo", PE: "Pérdida (gasto)", GA: "Ganancia (ingreso)" };
const TOL = 0.99; // tolerancia en pesos para descuadres por redondeo

/* ---------------- Utilidades ---------------- */
function norm(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function fmt(n) {
  if (n === null || n === undefined || n === "") return "";
  const r = Math.round(n * 10) / 10;
  return r.toLocaleString("es-CL", { maximumFractionDigits: 1 });
}

function parseNum(str) {
  if (typeof str === "number") return str;
  let s = String(str || "").replace(/[$\s]/g, "");
  if (!s) return 0;
  if (s.includes(".") && s.includes(",")) {
    s = s.replace(/\./g, "").replace(",", ".");
  } else if (s.includes(",")) {
    s = s.replace(",", ".");
  } else if (/^\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, "");
  }
  const v = parseFloat(s);
  return isNaN(v) ? 0 : v;
}

function fmtFecha(v) {
  if (v instanceof Date && !isNaN(v)) {
    const d = String(v.getDate()).padStart(2, "0");
    const m = String(v.getMonth() + 1).padStart(2, "0");
    return `${d}.${m}.${v.getFullYear()}`;
  }
  return String(v || "").trim();
}

function esFecha(v) {
  if (v instanceof Date && !isNaN(v)) return true;
  if (typeof v === "string" && /^\d{1,2}[./-]\d{1,2}([./-]\d{2,4})?$/.test(v.trim())) return true;
  return false;
}

/* ---------------- Clasificación de cuentas ---------------- */
const DICC = {
  A: ["caja", "banco", "bancos", "clientes", "mercaderia", "mercaderias", "materiales",
      "vehiculo", "vehiculos", "deposito a plazo", "iva credito fiscal", "arriendo anticipado",
      "maquinas y herramientas", "herramientas", "documentos por cobrar", "letras por cobrar",
      "deudores", "deudores varios", "muebles", "muebles y utiles", "terrenos", "edificios",
      "existencias", "efectivo", "equipos", "maquinarias", "utiles de oficina",
      "seguros anticipados", "gastos anticipados", "anticipo de sueldos", "inversiones"],
  P: ["letras por pagar", "documentos por pagar", "proveedores", "acreedores", "acreedores varios",
      "iva debito fiscal", "impuesto a pagar", "impuestos por pagar", "capital",
      "prestamo bancario", "prestamos bancarios", "remuneraciones por pagar", "sueldos por pagar",
      "retenciones", "provisiones", "utilidades acumuladas", "reservas", "cuentas por pagar"],
  PE: ["costo de venta", "costo de ventas", "gastos generales", "arriendo", "arriendos",
       "remuneraciones", "sueldos", "gastos financieros", "gastos de administracion",
       "depreciacion", "intereses pagados", "mantencion", "publicidad", "honorarios",
       "luz", "agua", "telefono", "combustible", "perdida"],
  GA: ["ingresos por ventas", "ventas", "ingresos", "intereses ganados", "intereses cobrados",
       "otros ingresos", "ingresos financieros", "utilidad en ventas"],
};

function clasificar(cuenta, saldoD, saldoA) {
  const n = norm(cuenta);
  if (clasifOverrides[n]) return { tipo: clasifOverrides[n], auto: false };
  for (const tipo of ["A", "P", "PE", "GA"]) {
    if (DICC[tipo].includes(n)) return { tipo, auto: false };
  }
  // Heurísticas por palabras clave
  if (n.includes("iva credito")) return { tipo: "A", auto: false };
  if (n.includes("iva debito")) return { tipo: "P", auto: false };
  if (n.includes("por pagar") || n.includes("capital") || n.includes("reserva")) return { tipo: "P", auto: false };
  if (n.includes("por cobrar") || n.includes("anticipado")) return { tipo: "A", auto: false };
  if (n.startsWith("costo") || n.startsWith("gasto") || n.includes("perdida")) return { tipo: "PE", auto: false };
  if (n.startsWith("ingreso") || n.includes("ganado") || n.includes("ganancia")) return { tipo: "GA", auto: false };
  // Sin coincidencia: se propone según el saldo y se pide confirmación al usuario
  return { tipo: saldoA > saldoD ? "P" : "A", auto: true };
}

/* ---------------- Lectura de Excel ---------------- */
function leerArchivo(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
      const res = parsearWorkbook(wb);
      if (!res.entries.length) {
        setAlert("alertas-diario", [{ tipo: "error", msg: "No se encontraron asientos contables en el archivo. Revisa que tenga el formato del ejercicio o de la plantilla (columnas Fecha, Cuenta, Debe, Haber)." }]);
        return;
      }
      entries = res.entries;
      alertaFuente = { tipo: "ok", msg: `Archivo "${file.name}" leído correctamente: ${entries.length} asientos detectados en la hoja "${res.hoja}".` };
      recalcularTodo();
    } catch (err) {
      setAlert("alertas-diario", [{ tipo: "error", msg: "Error al leer el archivo: " + err.message }]);
    }
  };
  reader.readAsArrayBuffer(file);
}

function parsearWorkbook(wb) {
  let mejor = { entries: [], hoja: "" };
  const ordenadas = [...wb.SheetNames].sort((a, b) => {
    const pa = norm(a).includes("diario") ? 0 : 1;
    const pb = norm(b).includes("diario") ? 0 : 1;
    return pa - pb;
  });
  for (const name of ordenadas) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    let res = parsearTabular(rows);
    if (!res.length) res = parsearFormatoEjercicio(rows);
    if (res.length > mejor.entries.length) mejor = { entries: res, hoja: name };
    if (mejor.entries.length && norm(name).includes("diario")) break;
  }
  return mejor;
}

/* Formato tabular (plantilla): encabezados Fecha | N° | Cuenta | Debe | Haber | Glosa */
function parsearTabular(rows) {
  let hIdx = -1, cols = {};
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const r = (rows[i] || []).map(norm);
    const ic = r.findIndex((v) => v === "cuenta" || v === "detalle" || v === "cuentas");
    const id = r.findIndex((v) => v === "debe" || v === "debitos" || v === "debito");
    const ih = r.findIndex((v) => v === "haber" || v === "creditos" || v === "credito");
    if (ic >= 0 && id >= 0 && ih >= 0) {
      hIdx = i;
      cols = {
        cuenta: ic, debe: id, haber: ih,
        fecha: r.findIndex((v) => v.startsWith("fecha")),
        num: r.findIndex((v) => v === "n°" || v === "nº" || v === "n" || v.startsWith("asiento") || v.startsWith("num")),
        glosa: r.findIndex((v) => v.startsWith("glosa") || v === "detalle glosa"),
      };
      break;
    }
  }
  if (hIdx < 0) return [];

  const out = [];
  let cur = null;
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const cuenta = r[cols.cuenta];
    if (norm(cuenta).startsWith("total")) break;
    const debe = parseNum(r[cols.debe] || 0);
    const haber = parseNum(r[cols.haber] || 0);
    const fecha = cols.fecha >= 0 ? r[cols.fecha] : null;
    const numV = cols.num >= 0 ? r[cols.num] : null;
    const glosa = cols.glosa >= 0 ? r[cols.glosa] : null;

    const nuevo =
      (numV !== null && numV !== "" && (!cur || String(numV) !== String(cur.num))) ||
      (cols.num < 0 && fecha !== null && fecha !== "" && cur && cur.lineas.length > 0);
    if (!cur || nuevo) {
      cur = { num: numV != null && numV !== "" ? numV : out.length + 1, fecha: fmtFecha(fecha), glosa: "", lineas: [] };
      out.push(cur);
    }
    if (fecha && !cur.fecha) cur.fecha = fmtFecha(fecha);
    if (glosa) cur.glosa = String(glosa).replace(/^glosa[:.]?\s*/i, "").trim();
    if (cuenta != null && String(cuenta).trim() !== "" && (debe || haber)) {
      cur.lineas.push({ cuenta: String(cuenta).trim(), debe, haber });
    }
  }
  return out.filter((e) => e.lineas.length);
}

/* Formato del ejercicio original: fecha y cuentas al Debe en una columna,
   cuentas al Haber desplazadas una columna, montos en dos columnas (Debe/Haber). */
function parsearFormatoEjercicio(rows) {
  // 1) columna de fechas = la que más celdas tipo fecha contiene
  const cuentaFechas = {};
  rows.forEach((r) => (r || []).forEach((v, c) => { if (esFecha(v)) cuentaFechas[c] = (cuentaFechas[c] || 0) + 1; }));
  const dateCol = Object.keys(cuentaFechas).length
    ? +Object.entries(cuentaFechas).sort((a, b) => b[1] - a[1])[0][0] : -1;
  if (dateCol < 0) return [];

  // 2) fila de término: "Totales"
  let endRow = rows.length;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    for (let c = Math.max(0, dateCol - 2); c <= dateCol + 3; c++) {
      if (norm(r[c]).startsWith("totales")) { endRow = i; break; }
    }
    if (endRow !== rows.length) break;
  }

  const firstDateRow = rows.findIndex((r) => esFecha((r || [])[dateCol]));
  if (firstDateRow < 0) return [];

  // 3) columnas de montos: las 2 con más números (cerca de la columna de fechas),
  //    contando solo filas de detalle (ni fecha ni glosa)
  const cuentaNums = {};
  for (let i = firstDateRow; i < endRow; i++) {
    const r = rows[i] || [];
    if (esFecha(r[dateCol])) continue;
    if (norm(r[dateCol]).startsWith("glosa") || norm(r[dateCol + 1]).startsWith("glosa")) continue;
    for (let c = dateCol + 1; c <= dateCol + 5; c++) {
      if (typeof r[c] === "number") cuentaNums[c] = (cuentaNums[c] || 0) + 1;
    }
  }
  const top = Object.entries(cuentaNums).sort((a, b) => b[1] - a[1]).slice(0, 2).map((e) => +e[0]).sort((a, b) => a - b);
  if (top.length < 2) return [];
  const [debeCol, haberCol] = top;

  // 4) recorrer y armar asientos
  const out = [];
  let cur = null;
  for (let i = firstDateRow; i < endRow; i++) {
    const r = rows[i] || [];
    if (esFecha(r[dateCol])) {
      let numV = null;
      for (let c = dateCol + 1; c < debeCol; c++) if (typeof r[c] === "number") { numV = r[c]; break; }
      cur = { num: numV != null ? numV : out.length + 1, fecha: fmtFecha(r[dateCol]), glosa: "", lineas: [] };
      out.push(cur);
      continue;
    }
    if (!cur) continue;
    const textoGlosa = [r[dateCol], r[dateCol + 1]].find((v) => norm(v).startsWith("glosa"));
    if (textoGlosa) { cur.glosa = String(textoGlosa).replace(/^glosa[:.]?\s*/i, "").trim(); continue; }

    const buscarCuenta = (hasta) => {
      for (let c = dateCol; c < hasta; c++) {
        const v = r[c];
        if (typeof v === "string" && v.trim() !== "") return v.trim();
      }
      return null;
    };
    if (typeof r[debeCol] === "number" && r[debeCol] !== 0) {
      const cta = buscarCuenta(debeCol);
      if (cta) cur.lineas.push({ cuenta: cta, debe: r[debeCol], haber: 0 });
    }
    if (typeof r[haberCol] === "number" && r[haberCol] !== 0) {
      const cta = buscarCuenta(haberCol);
      if (cta) cur.lineas.push({ cuenta: cta, debe: 0, haber: r[haberCol] });
    }
  }
  return out.filter((e) => e.lineas.length);
}

/* ---------------- Cálculo: Libro Mayor ---------------- */
function computarMayor() {
  const mapa = new Map(); // norm -> {nombre, movs:[{num, debe, haber}], sumD, sumH, saldoD, saldoA}
  for (const e of entries) {
    for (const l of e.lineas) {
      const n = norm(l.cuenta);
      if (!n) continue;
      if (!mapa.has(n)) mapa.set(n, { nombre: l.cuenta.trim(), movs: [], sumD: 0, sumH: 0 });
      const cta = mapa.get(n);
      cta.movs.push({ num: e.num, fecha: e.fecha, debe: l.debe || 0, haber: l.haber || 0 });
      cta.sumD += l.debe || 0;
      cta.sumH += l.haber || 0;
    }
  }
  for (const cta of mapa.values()) {
    const dif = cta.sumD - cta.sumH;
    cta.saldoD = dif > TOL ? dif : 0;
    cta.saldoA = dif < -TOL ? -dif : 0;
  }
  return mapa;
}

/* ---------------- Cálculo: Balance 8 columnas ---------------- */
function computarBalance(mayor) {
  const filas = [];
  for (const cta of mayor.values()) {
    const { tipo, auto } = clasificar(cta.nombre, cta.saldoD, cta.saldoA);
    const f = {
      nombre: cta.nombre, tipo, auto,
      debitos: cta.sumD, creditos: cta.sumH,
      saldoD: cta.saldoD, saldoA: cta.saldoA,
      activo: 0, pasivo: 0, perdida: 0, ganancia: 0,
    };
    if (tipo === "A") f.activo = f.saldoD - f.saldoA;
    else if (tipo === "P") f.pasivo = f.saldoA - f.saldoD;
    else if (tipo === "PE") f.perdida = f.saldoD - f.saldoA;
    else f.ganancia = f.saldoA - f.saldoD;
    filas.push(f);
  }
  const sub = filas.reduce((a, f) => ({
    debitos: a.debitos + f.debitos, creditos: a.creditos + f.creditos,
    saldoD: a.saldoD + f.saldoD, saldoA: a.saldoA + f.saldoA,
    activo: a.activo + f.activo, pasivo: a.pasivo + f.pasivo,
    perdida: a.perdida + f.perdida, ganancia: a.ganancia + f.ganancia,
  }), { debitos: 0, creditos: 0, saldoD: 0, saldoA: 0, activo: 0, pasivo: 0, perdida: 0, ganancia: 0 });

  const resultado = sub.ganancia - sub.perdida; // >0 utilidad, <0 pérdida
  const tot = {
    debitos: sub.debitos, creditos: sub.creditos, saldoD: sub.saldoD, saldoA: sub.saldoA,
    activo: sub.activo + (resultado < 0 ? -resultado : 0),
    pasivo: sub.pasivo + (resultado > 0 ? resultado : 0),
    perdida: sub.perdida + (resultado > 0 ? resultado : 0),
    ganancia: sub.ganancia + (resultado < 0 ? -resultado : 0),
  };
  return { filas, sub, resultado, tot };
}

/* ---------------- Validaciones ---------------- */
function validarDiario() {
  const alerts = [];
  let totD = 0, totH = 0;
  const descuadrados = [];
  for (const e of entries) {
    const d = e.lineas.reduce((a, l) => a + (l.debe || 0), 0);
    const h = e.lineas.reduce((a, l) => a + (l.haber || 0), 0);
    totD += d; totH += h;
    if (Math.abs(d - h) > TOL) descuadrados.push(`N° ${e.num} (diferencia ${fmt(d - h)})`);
  }
  if (descuadrados.length) {
    alerts.push({ tipo: "error", msg: `⚠ Asientos descuadrados (Debe ≠ Haber): ${descuadrados.join(", ")}. Corrígelos antes de usar el Mayor y el Balance.` });
  }
  if (Math.abs(totD - totH) <= TOL && entries.length) {
    alerts.push({ tipo: "ok", msg: `✔ Libro Diario cuadrado: total Debe = total Haber = $ ${fmt(totD)} (partida doble verificada en los ${entries.length} asientos).` });
  } else if (entries.length) {
    alerts.push({ tipo: "error", msg: `✖ El Libro Diario NO cuadra: Debe $ ${fmt(totD)} vs Haber $ ${fmt(totH)} (diferencia $ ${fmt(totD - totH)}).` });
  }
  return { alerts, totD, totH, ok: Math.abs(totD - totH) <= TOL && !descuadrados.length };
}

/* ---------------- Render: Diario editable ---------------- */
function renderDiario(val) {
  const wrap = document.getElementById("diario-wrap");
  if (!entries.length) { wrap.classList.add("hidden"); return; }
  wrap.classList.remove("hidden");

  let html = `<table class="grid"><thead><tr>
    <th style="width:36px"></th><th>N°</th><th>Fecha</th><th>Cuenta</th>
    <th class="num">Debe</th><th class="num">Haber</th><th style="width:36px"></th>
  </tr></thead><tbody>`;

  entries.forEach((e, ei) => {
    const d = e.lineas.reduce((a, l) => a + (l.debe || 0), 0);
    const h = e.lineas.reduce((a, l) => a + (l.haber || 0), 0);
    const ok = Math.abs(d - h) <= TOL;
    html += `<tr class="asiento-sep">
      <td><button class="mini-btn" title="Eliminar asiento" onclick="eliminarAsiento(${ei})">🗑</button></td>
      <td><input class="cell" style="width:52px" value="${escapeHtml(e.num)}" onchange="editarAsiento(${ei},'num',this.value)"></td>
      <td><input class="cell" style="width:110px" value="${escapeHtml(e.fecha)}" onchange="editarAsiento(${ei},'fecha',this.value)"></td>
      <td colspan="2">Asiento ${escapeHtml(e.num)} — ${d === h ? "" : ""}<span class="asiento-check ${ok ? "ok" : "bad"}">${ok ? "✔ cuadrado" : `✖ descuadrado ($ ${fmt(d - h)})`}</span></td>
      <td class="num">${fmt(d)} / ${fmt(h)}</td>
      <td><button class="mini-btn add" title="Agregar línea" onclick="agregarLinea(${ei})">＋</button></td>
    </tr>`;
    e.lineas.forEach((l, li) => {
      const esHaber = !l.debe && l.haber;
      html += `<tr>
        <td></td><td></td><td></td>
        <td style="${esHaber ? "padding-left:42px" : ""}">
          <input class="cell" value="${escapeHtml(l.cuenta)}" onchange="editarLinea(${ei},${li},'cuenta',this.value)">
        </td>
        <td class="num"><input class="cell-num" value="${l.debe ? fmt(l.debe) : ""}" onchange="editarLinea(${ei},${li},'debe',this.value)"></td>
        <td class="num"><input class="cell-num" value="${l.haber ? fmt(l.haber) : ""}" onchange="editarLinea(${ei},${li},'haber',this.value)"></td>
        <td><button class="mini-btn" title="Eliminar línea" onclick="eliminarLinea(${ei},${li})">✕</button></td>
      </tr>`;
    });
    html += `<tr class="glosa-row"><td></td><td></td><td></td>
      <td colspan="3"><input class="cell" placeholder="Glosa…" value="${escapeHtml(e.glosa)}" onchange="editarAsiento(${ei},'glosa',this.value)"></td><td></td></tr>`;
  });
  html += `</tbody></table>`;
  document.getElementById("diario-tabla").innerHTML = html;

  document.getElementById("diario-resumen").textContent =
    `${entries.length} asientos · ${entries.reduce((a, e) => a + e.lineas.length, 0)} líneas`;
  document.getElementById("diario-totales").innerHTML = `
    <span><span class="tag">Total Debe:</span>$ ${fmt(val.totD)}</span>
    <span><span class="tag">Total Haber:</span>$ ${fmt(val.totH)}</span>
    <span style="color:${val.ok ? "var(--accent)" : "var(--danger)"}">${val.ok ? "✔ PARTIDA DOBLE OK" : "✖ REVISAR DESCUADRE"}</span>`;
}

function escapeHtml(s) {
  return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
}

/* Ediciones desde la tabla (recalculan todo al instante) */
window.editarAsiento = (ei, campo, valor) => { entries[ei][campo] = valor; recalcularTodo(); };
window.editarLinea = (ei, li, campo, valor) => {
  const l = entries[ei].lineas[li];
  if (campo === "cuenta") l.cuenta = valor;
  else l[campo] = parseNum(valor);
  recalcularTodo();
};
window.eliminarLinea = (ei, li) => { entries[ei].lineas.splice(li, 1); recalcularTodo(); };
window.eliminarAsiento = (ei) => { entries.splice(ei, 1); recalcularTodo(); };
window.agregarLinea = (ei) => { entries[ei].lineas.push({ cuenta: "", debe: 0, haber: 0 }); recalcularTodo(); };
window.cambiarClasif = (nombreNorm, tipo) => { clasifOverrides[nombreNorm] = tipo; recalcularTodo(); };

/* ---------------- Render: Mayor ---------------- */
function renderMayor(mayor) {
  const grid = document.getElementById("mayor-grid");
  let html = "";
  for (const cta of mayor.values()) {
    const nMovs = Math.max(cta.movs.filter((m) => m.debe).length, cta.movs.filter((m) => m.haber).length, 1);
    const debes = cta.movs.filter((m) => m.debe);
    const habers = cta.movs.filter((m) => m.haber);
    let filas = "";
    for (let i = 0; i < nMovs; i++) {
      filas += `<tr>
        <td class="num">${debes[i] ? fmt(debes[i].debe) : ""}</td>
        <td class="num">${habers[i] ? fmt(habers[i].haber) : ""}</td>
      </tr>`;
    }
    const saldoTag = cta.saldoD > 0
      ? `<span class="saldo-tag sd">SD $ ${fmt(cta.saldoD)}</span>`
      : cta.saldoA > 0
        ? `<span class="saldo-tag sa">SA $ ${fmt(cta.saldoA)}</span>`
        : `<span class="saldo-tag cero">Saldada</span>`;
    html += `<div class="tcard">
      <div class="tcard-head"><span>${escapeHtml(cta.nombre)}</span>${saldoTag}</div>
      <table>
        <tr><th style="width:50%">DEBE</th><th>HABER</th></tr>
        ${filas}
        <tr class="sums"><td class="num">${fmt(cta.sumD)}</td><td class="num">${fmt(cta.sumH)}</td></tr>
        <tr class="saldo"><td colspan="2">Saldo: ${cta.saldoD > 0 ? "Deudor (SD) $ " + fmt(cta.saldoD) : cta.saldoA > 0 ? "Acreedor (SA) $ " + fmt(cta.saldoA) : "$ 0 — cuenta saldada"}</td></tr>
      </table>
    </div>`;
  }
  grid.innerHTML = html;
  document.getElementById("mayor-resumen").textContent = `${mayor.size} cuentas`;
  setAlert("alertas-mayor", entries.length ? [{
    tipo: "ok",
    msg: `✔ Mayorización automática: ${entries.reduce((a, e) => a + e.lineas.length, 0)} movimientos del Diario traspasados a ${mayor.size} cuentas, agrupados y con saldos calculados sin intervención manual.`,
  }] : []);
}

/* ---------------- Render: Balance ---------------- */
function renderBalance(bal) {
  const { filas, sub, resultado, tot } = bal;
  const guion = `<td class="num">–</td>`;
  const celda = (v) => `<td class="num">${fmt(v)}</td>`;
  const celdaOG = (v, cond) => (cond ? celda(v) : guion);

  let html = `<table class="grid"><thead><tr>
    <th>CUENTAS</th><th style="width:130px">Clasificación</th>
    <th class="num">DÉBITOS</th><th class="num">CRÉDITOS</th>
    <th class="num">SALDO DEUDOR</th><th class="num">SALDO ACREEDOR</th>
    <th class="num">ACTIVOS</th><th class="num">PASIVOS</th>
    <th class="num">PÉRDIDAS</th><th class="num">GANANCIAS</th>
  </tr></thead><tbody>`;

  const sinConfirmar = [];
  for (const f of filas) {
    if (f.auto) sinConfirmar.push(f.nombre);
    const n = norm(f.nombre);
    html += `<tr>
      <td>${escapeHtml(f.nombre)}</td>
      <td><select class="clasif ${f.auto ? "auto-warn" : ""}" onchange="cambiarClasif('${n.replace(/'/g, "\\'")}', this.value)">
        ${["A", "P", "PE", "GA"].map((t) => `<option value="${t}" ${t === f.tipo ? "selected" : ""}>${TIPOS[t]}</option>`).join("")}
      </select></td>
      ${celda(f.debitos)}${celda(f.creditos)}
      ${celdaOG(f.saldoD, f.saldoD > 0 || f.saldoA === 0)}${celdaOG(f.saldoA, f.saldoA > 0)}
      ${f.tipo === "A" ? celda(f.activo) : guion}
      ${f.tipo === "P" ? celda(f.pasivo) : guion}
      ${f.tipo === "PE" ? celda(f.perdida) : guion}
      ${f.tipo === "GA" ? celda(f.ganancia) : guion}
    </tr>`;
  }

  html += `<tr class="subtotal"><td>Subtotal</td><td></td>
    ${celda(sub.debitos)}${celda(sub.creditos)}${celda(sub.saldoD)}${celda(sub.saldoA)}
    ${celda(sub.activo)}${celda(sub.pasivo)}${celda(sub.perdida)}${celda(sub.ganancia)}</tr>`;

  const esUtil = resultado >= 0;
  html += `<tr class="utilidad ${esUtil ? "" : "perdida"}">
    <td>${esUtil ? "Utilidad del Ejercicio" : "Pérdida del Ejercicio"}</td>
    <td></td><td colspan="4"></td>
    <td class="num">${esUtil ? "" : fmt(-resultado)}</td>
    <td class="num">${esUtil ? fmt(resultado) : ""}</td>
    <td class="num">${esUtil ? fmt(resultado) : ""}</td>
    <td class="num">${esUtil ? "" : fmt(-resultado)}</td>
  </tr>`;

  html += `<tr class="total"><td>TOTALES IGUALES</td><td></td>
    ${celda(tot.debitos)}${celda(tot.creditos)}${celda(tot.saldoD)}${celda(tot.saldoA)}
    ${celda(tot.activo)}${celda(tot.pasivo)}${celda(tot.perdida)}${celda(tot.ganancia)}</tr>`;
  html += `</tbody></table>`;
  document.getElementById("balance-tabla").innerHTML = html;

  // Validaciones del balance
  const checks = [
    { ok: Math.abs(tot.debitos - tot.creditos) <= TOL, txt: "Débitos = Créditos" },
    { ok: Math.abs(tot.saldoD - tot.saldoA) <= TOL, txt: "Saldo Deudor = Saldo Acreedor" },
    { ok: Math.abs(tot.activo - tot.pasivo) <= TOL, txt: "Activos = Pasivos + Resultado" },
    { ok: Math.abs(tot.perdida - tot.ganancia) <= TOL, txt: "Pérdidas = Ganancias (con resultado)" },
  ];
  const todoOk = checks.every((c) => c.ok);
  const estado = document.getElementById("balance-estado");
  estado.textContent = todoOk ? "✔ BALANCE CUADRADO" : "✖ BALANCE DESCUADRADO";
  estado.className = "pill " + (todoOk ? "ok" : "bad");

  const alerts = [];
  alerts.push(todoOk
    ? { tipo: "ok", msg: `✔ BALANCE CUADRADO: ${checks.map((c) => c.txt).join(" · ")}. ${resultado >= 0 ? "Utilidad" : "Pérdida"} del ejercicio: $ ${fmt(Math.abs(resultado))}.` }
    : { tipo: "error", msg: `✖ BALANCE DESCUADRADO: falla ${checks.filter((c) => !c.ok).map((c) => c.txt).join(", ")}. Revisa los asientos marcados en el Libro Diario.` });
  if (sinConfirmar.length) {
    alerts.push({ tipo: "warn", msg: `⚠ Cuentas clasificadas automáticamente por su saldo (confirma su tipo con el selector): ${sinConfirmar.join(", ")}.` });
  }
  setAlert("alertas-balance", alerts);
}

/* ---------------- Render: Estado de Resultados ---------------- */
function renderEERR(bal) {
  const ingresos = bal.filas.filter((f) => f.tipo === "GA");
  const gastos = bal.filas.filter((f) => f.tipo === "PE");
  const ti = ingresos.reduce((a, f) => a + f.ganancia, 0);
  const tg = gastos.reduce((a, f) => a + f.perdida, 0);
  const res = ti - tg;
  let html = `<table class="grid"><thead><tr><th>ESTADO DE RESULTADOS</th><th class="num">Monto ($)</th></tr></thead><tbody>`;
  html += `<tr class="subtotal"><td>INGRESOS</td><td></td></tr>`;
  ingresos.forEach((f) => { html += `<tr><td style="padding-left:26px">${escapeHtml(f.nombre)}</td><td class="num">${fmt(f.ganancia)}</td></tr>`; });
  html += `<tr class="subtotal"><td>Total Ingresos</td><td class="num">${fmt(ti)}</td></tr>`;
  html += `<tr class="subtotal"><td>GASTOS</td><td></td></tr>`;
  gastos.forEach((f) => { html += `<tr><td style="padding-left:26px">${escapeHtml(f.nombre)}</td><td class="num">(${fmt(f.perdida)})</td></tr>`; });
  html += `<tr class="subtotal"><td>Total Gastos</td><td class="num">(${fmt(tg)})</td></tr>`;
  html += `<tr class="total utilidad ${res < 0 ? "perdida" : ""}"><td>${res >= 0 ? "UTILIDAD DEL EJERCICIO" : "PÉRDIDA DEL EJERCICIO"}</td><td class="num">${fmt(res)}</td></tr>`;
  html += `</tbody></table>`;
  document.getElementById("eerr-tabla").innerHTML = html;
}

/* ---------------- Alertas ---------------- */
function setAlert(id, list) {
  document.getElementById(id).innerHTML = list
    .map((a) => `<div class="alert ${a.tipo}">${a.msg}</div>`).join("");
}

/* ---------------- Recalcular todo ---------------- */
let _mayor = new Map(), _balance = null;
function recalcularTodo() {
  const val = validarDiario();
  renderDiario(val);
  setAlert("alertas-diario", [alertaFuente, ...val.alerts].filter(Boolean));
  _mayor = computarMayor();
  _balance = computarBalance(_mayor);
  renderMayor(_mayor);
  renderBalance(_balance);
  renderEERR(_balance);
  const hay = entries.length > 0;
  ["nav-mayor", "nav-balance", "nav-eerr"].forEach((id) => (document.getElementById(id).disabled = !hay));
}

/* =====================================================================
   EXPORTACIÓN A EXCEL (con estilos, usando xlsx-js-style)
   ===================================================================== */
const S = {
  titulo: { font: { bold: true, sz: 14, color: { rgb: "26325E" } } },
  th: { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "26325E" } },
        alignment: { horizontal: "center" }, border: bordes() },
  num: { numFmt: "#,##0", border: bordes(), alignment: { horizontal: "right" } },
  txt: { border: bordes() },
  glosa: { font: { italic: true, color: { rgb: "5B6478" } }, border: bordes() },
  totRow: { font: { bold: true }, fill: { fgColor: { rgb: "EEF2FB" } }, numFmt: "#,##0", border: bordes("medium") },
  totTxt: { font: { bold: true }, fill: { fgColor: { rgb: "EEF2FB" } }, border: bordes("medium") },
  utilRow: { font: { bold: true }, fill: { fgColor: { rgb: "E3F7EF" } }, numFmt: "#,##0", border: bordes() },
  ok: { font: { bold: true, color: { rgb: "0A9D6C" } } },
  bad: { font: { bold: true, color: { rgb: "D23F3F" } } },
};
function bordes(st = "thin") {
  const b = { style: st, color: { rgb: "B9C6E4" } };
  return { top: b, bottom: b, left: b, right: b };
}
const C = {
  s: (v, st) => ({ v: v ?? "", t: "s", s: st || S.txt }),
  n: (v, st) => ({ v: Math.round((v ?? 0) * 100) / 100, t: "n", s: st || S.num }),
};

function hojaDiario() {
  const aoa = [
    [C.s("LIBRO DIARIO", S.titulo)],
    [],
    [C.s("N°", S.th), C.s("Fecha", S.th), C.s("Cuenta", S.th), C.s("Debe", S.th), C.s("Haber", S.th), C.s("Glosa", S.th)],
  ];
  let totD = 0, totH = 0;
  for (const e of entries) {
    e.lineas.forEach((l, i) => {
      totD += l.debe || 0; totH += l.haber || 0;
      aoa.push([
        C.s(i === 0 ? String(e.num) : ""),
        C.s(i === 0 ? e.fecha : ""),
        C.s((l.haber && !l.debe ? "        " : "") + l.cuenta),
        l.debe ? C.n(l.debe) : C.s("", S.num),
        l.haber ? C.n(l.haber) : C.s("", S.num),
        C.s(i === 0 ? "" : "", S.glosa),
      ]);
    });
    aoa.push([C.s(""), C.s(""), C.s("Glosa: " + (e.glosa || "—"), S.glosa), C.s("", S.glosa), C.s("", S.glosa), C.s("", S.glosa)]);
  }
  aoa.push([C.s("", S.totTxt), C.s("", S.totTxt), C.s("TOTALES", S.totTxt), C.n(totD, S.totRow), C.n(totH, S.totRow), C.s("", S.totTxt)]);
  aoa.push([]);
  const cuadrado = Math.abs(totD - totH) <= TOL;
  aoa.push([C.s(cuadrado ? "✔ LIBRO DIARIO CUADRADO (Debe = Haber)" : "✖ LIBRO DIARIO DESCUADRADO — revisar asientos", cuadrado ? S.ok : S.bad)]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 6 }, { wch: 12 }, { wch: 38 }, { wch: 15 }, { wch: 15 }, { wch: 50 }];
  return ws;
}

function hojaMayor() {
  const cuentas = [..._mayor.values()];
  const PORFILA = 3, ANCHO = 3; // 2 columnas + 1 separador
  const aoa = [[C.s("LIBRO MAYOR (Cuentas T)", S.titulo)], []];
  const merges = [];
  let base = 2;
  for (let g = 0; g < cuentas.length; g += PORFILA) {
    const grupo = cuentas.slice(g, g + PORFILA);
    const altura = Math.max(...grupo.map((c) =>
      Math.max(c.movs.filter((m) => m.debe).length, c.movs.filter((m) => m.haber).length))) + 4;
    for (let r = 0; r < altura + 1; r++) aoa.push([]);
    grupo.forEach((cta, k) => {
      const col = k * ANCHO;
      const fila = base;
      setCell(aoa, fila, col, C.s(cta.nombre, S.th));
      setCell(aoa, fila, col + 1, C.s("", S.th));
      merges.push({ s: { r: fila, c: col }, e: { r: fila, c: col + 1 } });
      setCell(aoa, fila + 1, col, C.s("DEBE", S.th));
      setCell(aoa, fila + 1, col + 1, C.s("HABER", S.th));
      const debes = cta.movs.filter((m) => m.debe), habers = cta.movs.filter((m) => m.haber);
      const nMovs = Math.max(debes.length, habers.length);
      for (let i = 0; i < nMovs; i++) {
        setCell(aoa, fila + 2 + i, col, debes[i] ? C.n(debes[i].debe) : C.s("", S.num));
        setCell(aoa, fila + 2 + i, col + 1, habers[i] ? C.n(habers[i].haber) : C.s("", S.num));
      }
      setCell(aoa, fila + 2 + nMovs, col, C.n(cta.sumD, S.totRow));
      setCell(aoa, fila + 2 + nMovs, col + 1, C.n(cta.sumH, S.totRow));
      const saldoTxt = cta.saldoD > 0 ? "Saldo Deudor (SD)" : cta.saldoA > 0 ? "Saldo Acreedor (SA)" : "Cuenta saldada";
      setCell(aoa, fila + 3 + nMovs, col, C.s(saldoTxt, S.utilRow));
      setCell(aoa, fila + 3 + nMovs, col + 1, C.n(cta.saldoD || cta.saldoA, S.utilRow));
    });
    base += altura + 1;
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = merges;
  ws["!cols"] = Array.from({ length: PORFILA * ANCHO }, (_, i) => ({ wch: i % ANCHO === 2 ? 3 : 16 }));
  return ws;
}

function setCell(aoa, r, c, cell) {
  while (aoa.length <= r) aoa.push([]);
  aoa[r][c] = cell;
}

function hojaBalance() {
  const { filas, sub, resultado, tot } = _balance;
  const g = () => C.s("–", { ...S.num, alignment: { horizontal: "center" } });
  const aoa = [
    [C.s("BALANCE DE 8 COLUMNAS", S.titulo)],
    [C.s("Generado automáticamente desde el Libro Mayor", S.glosa)],
    [],
    ["CUENTAS", "DÉBITOS", "CRÉDITOS", "SALDO DEUDOR", "SALDO ACREEDOR", "ACTIVOS", "PASIVOS", "PÉRDIDAS", "GANANCIAS"].map((h) => C.s(h, S.th)),
  ];
  for (const f of filas) {
    aoa.push([
      C.s(f.nombre), C.n(f.debitos), C.n(f.creditos),
      f.saldoD > 0 || f.saldoA === 0 ? C.n(f.saldoD) : g(),
      f.saldoA > 0 ? C.n(f.saldoA) : g(),
      f.tipo === "A" ? C.n(f.activo) : g(),
      f.tipo === "P" ? C.n(f.pasivo) : g(),
      f.tipo === "PE" ? C.n(f.perdida) : g(),
      f.tipo === "GA" ? C.n(f.ganancia) : g(),
    ]);
  }
  aoa.push([C.s("Subtotal", S.totTxt), C.n(sub.debitos, S.totRow), C.n(sub.creditos, S.totRow),
    C.n(sub.saldoD, S.totRow), C.n(sub.saldoA, S.totRow), C.n(sub.activo, S.totRow),
    C.n(sub.pasivo, S.totRow), C.n(sub.perdida, S.totRow), C.n(sub.ganancia, S.totRow)]);
  const esU = resultado >= 0;
  aoa.push([C.s(esU ? "Utilidad del Ejercicio" : "Pérdida del Ejercicio", S.utilRow),
    C.s("", S.utilRow), C.s("", S.utilRow), C.s("", S.utilRow), C.s("", S.utilRow),
    esU ? C.s("", S.utilRow) : C.n(-resultado, S.utilRow),
    esU ? C.n(resultado, S.utilRow) : C.s("", S.utilRow),
    esU ? C.n(resultado, S.utilRow) : C.s("", S.utilRow),
    esU ? C.s("", S.utilRow) : C.n(-resultado, S.utilRow)]);
  aoa.push([C.s("TOTALES IGUALES", S.totTxt), C.n(tot.debitos, S.totRow), C.n(tot.creditos, S.totRow),
    C.n(tot.saldoD, S.totRow), C.n(tot.saldoA, S.totRow), C.n(tot.activo, S.totRow),
    C.n(tot.pasivo, S.totRow), C.n(tot.perdida, S.totRow), C.n(tot.ganancia, S.totRow)]);
  aoa.push([]);
  const ok = Math.abs(tot.activo - tot.pasivo) <= TOL && Math.abs(tot.perdida - tot.ganancia) <= TOL;
  aoa.push([C.s(ok ? "✔ BALANCE CUADRADO" : "✖ BALANCE DESCUADRADO", ok ? S.ok : S.bad)]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 28 }, ...Array.from({ length: 8 }, () => ({ wch: 15 }))];
  return ws;
}

function hojaEERR() {
  const ingresos = _balance.filas.filter((f) => f.tipo === "GA");
  const gastos = _balance.filas.filter((f) => f.tipo === "PE");
  const ti = ingresos.reduce((a, f) => a + f.ganancia, 0);
  const tg = gastos.reduce((a, f) => a + f.perdida, 0);
  const aoa = [
    [C.s("ESTADO DE RESULTADOS", S.titulo)], [],
    [C.s("Concepto", S.th), C.s("Monto ($)", S.th)],
    [C.s("INGRESOS", S.totTxt), C.s("", S.totRow)],
    ...ingresos.map((f) => [C.s("   " + f.nombre), C.n(f.ganancia)]),
    [C.s("Total Ingresos", S.totTxt), C.n(ti, S.totRow)],
    [C.s("GASTOS", S.totTxt), C.s("", S.totRow)],
    ...gastos.map((f) => [C.s("   " + f.nombre), C.n(-f.perdida)]),
    [C.s("Total Gastos", S.totTxt), C.n(-tg, S.totRow)],
    [C.s(ti - tg >= 0 ? "UTILIDAD DEL EJERCICIO" : "PÉRDIDA DEL EJERCICIO", S.totTxt), C.n(ti - tg, S.utilRow)],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 32 }, { wch: 16 }];
  return ws;
}

function exportar(hojas, nombre) {
  if (!entries.length) return;
  const wb = XLSX.utils.book_new();
  for (const [nombreHoja, ws] of hojas) XLSX.utils.book_append_sheet(wb, ws, nombreHoja);
  XLSX.writeFile(wb, nombre);
}

function descargarPlantilla() {
  const aoa = [
    [C.s("PLANTILLA LIBRO DIARIO — completa tus asientos y súbela a la página", S.titulo)],
    [C.s("Cada asiento: mismo N° en todas sus líneas. Debe = Haber por asiento.", S.glosa)],
    [],
    [C.s("N°", S.th), C.s("Fecha", S.th), C.s("Cuenta", S.th), C.s("Debe", S.th), C.s("Haber", S.th), C.s("Glosa", S.th)],
    [C.n(1, S.txt), C.s("01.02.2026"), C.s("Caja"), C.n(1000000), C.n(0), C.s("Asiento de apertura", S.glosa)],
    [C.n(1, S.txt), C.s("01.02.2026"), C.s("Capital"), C.n(0), C.n(1000000), C.s("", S.glosa)],
    [C.n(2, S.txt), C.s("05.02.2026"), C.s("Banco"), C.n(500000), C.n(0), C.s("Apertura cuenta corriente", S.glosa)],
    [C.n(2, S.txt), C.s("05.02.2026"), C.s("Caja"), C.n(0), C.n(500000), C.s("", S.glosa)],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 6 }, { wch: 12 }, { wch: 30 }, { wch: 14 }, { wch: 14 }, { wch: 40 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Libro Diario");
  XLSX.writeFile(wb, "Plantilla_Libro_Diario.xlsx");
}

/* ---------------- Navegación y eventos ---------------- */
document.querySelectorAll(".step-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    if (btn.disabled) return;
    document.querySelectorAll(".step-btn").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".panel").forEach((p) => p.classList.remove("visible"));
    btn.classList.add("active");
    document.getElementById(btn.dataset.target).classList.add("visible");
  });
});

const dz = document.getElementById("dropzone");
const fi = document.getElementById("file-input");
fi.addEventListener("change", () => { if (fi.files[0]) leerArchivo(fi.files[0]); fi.value = ""; });
["dragover", "dragenter"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("dragover"); }));
["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("dragover"); }));
dz.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) leerArchivo(f); });

document.getElementById("btn-plantilla").addEventListener("click", descargarPlantilla);
document.getElementById("btn-add-asiento").addEventListener("click", () => {
  entries.push({
    num: entries.length + 1, fecha: "", glosa: "",
    lineas: [{ cuenta: "", debe: 0, haber: 0 }, { cuenta: "", debe: 0, haber: 0 }],
  });
  recalcularTodo();
});
document.getElementById("btn-limpiar").addEventListener("click", () => {
  entries = [];
  alertaFuente = null;
  ["alertas-diario", "alertas-mayor", "alertas-balance"].forEach((id) => setAlert(id, []));
  document.getElementById("diario-wrap").classList.add("hidden");
  ["mayor-grid", "balance-tabla", "eerr-tabla"].forEach((id) => (document.getElementById(id).innerHTML = ""));
  ["nav-mayor", "nav-balance", "nav-eerr"].forEach((id) => (document.getElementById(id).disabled = true));
  document.querySelector('[data-target="sec-diario"]').click();
});

document.getElementById("btn-export-diario").addEventListener("click", () =>
  exportar([["Libro Diario", hojaDiario()]], "Libro_Diario.xlsx"));
document.getElementById("btn-export-mayor").addEventListener("click", () =>
  exportar([["Libro Mayor", hojaMayor()]], "Libro_Mayor.xlsx"));
document.getElementById("btn-export-diario-mayor").addEventListener("click", () =>
  exportar([["Libro Diario", hojaDiario()], ["Libro Mayor", hojaMayor()]], "Libro_Diario_y_Mayor.xlsx"));
document.getElementById("btn-export-balance").addEventListener("click", () =>
  exportar([["Balance 8 Columnas", hojaBalance()]], "Balance_8_Columnas.xlsx"));
document.getElementById("btn-export-todo").addEventListener("click", () =>
  exportar([
    ["Libro Diario", hojaDiario()],
    ["Libro Mayor", hojaMayor()],
    ["Balance 8 Columnas", hojaBalance()],
    ["Estado de Resultados", hojaEERR()],
  ], "Ciclo_Contable_Completo.xlsx"));
