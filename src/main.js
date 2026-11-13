import "./style.css";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import { pageText, parsePoliza, parseFacturaLines } from "./parsers.js";

// El lector de PDF se carga solo cuando se suben archivos.
let pdfjs = null;
async function loadPdfjs() {
  if (!pdfjs) {
    const [lib, worker] = await Promise.all([import("pdfjs-dist"), import("pdfjs-dist/build/pdf.worker.min.mjs?url")]);
    lib.GlobalWorkerOptions.workerSrc = worker.default;
    pdfjs = lib;
  }
  return pdfjs;
}

/* ================= Supabase ================= */
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_KEY = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || import.meta.env.VITE_SUPABASE_ANON_KEY;
// El correo del usuario viaja en el encabezado x-usuario; la base solo
// muestra datos y acepta cambios si ese correo está en seg_usuarios.
let sb = null;
function conectar(email) {
  sb = createClient(SUPABASE_URL, SUPABASE_KEY, {
    auth: { persistSession: false, autoRefreshToken: false },
    global: { headers: email ? { "x-usuario": email } : {} }
  });
}

async function fetchAll(build) {
  const out = [], size = 1000;
  for (let from = 0; ; from += size) {
    const { data, error } = await build().range(from, from + size - 1);
    if (error) throw error;
    out.push(...data);
    if (data.length < size) return out;
  }
}
async function rpc(fn, args) {
  const { data, error } = await sb.rpc(fn, args);
  if (error) {
    if (/Usuario no habilitado/.test(error.message || "")) salir("Tu correo ya no está habilitado.");
    throw error;
  }
  return data;
}
const errMsg = e => (e && (e.message || e.error_description)) || "Error desconocido";

/* ================= Helpers ================= */
const COMP = ["neto", "iva", "perc", "otros"];
const fmt = n => (n || 0).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtPct = n => (n * 100).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " %";
const fmtInt = n => n.toLocaleString("es-AR");
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const normH = s => String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
const normPat = s => String(s ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "");
const isoToAr = d => (d ? d.split("-").reverse().join("/") : "");
const $ = id => document.getElementById(id);
let toastT;
function toast(t) { const el = $("toast"); el.textContent = t; el.classList.add("on"); clearTimeout(toastT); toastT = setTimeout(() => el.classList.remove("on"), 4000); }
function busy(btn, on, label) { btn.disabled = on; btn.classList.toggle("busy", on); if (label) btn.textContent = label; }

/* ================= State ================= */
const S = {
  periodos: [],           // filas de seg_periodos
  actual: null,           // id del período visible
  ccBase: [],             // [{codigo, descripcion, habilitado}]
  bases: {},              // tipo -> {archivo, filas, detalle, cargado_en}
  zonas: {},              // patente -> zona
  asignadas: new Set(),   // patentes con CC recordado
  view: "dash", dashAll: false, q: "", fil: "all", zf: "", bulkCC: "", sel: new Set(), sort: { k: "cert", dir: 1 }, open: new Set(),
  up: { poliza: null, factura: null, cc: null, zona: null },
  pdfs: {}                // periodo -> datos del PDF guardado de la factura (null si falta la migración)
};
let ccMap = new Map();
let ccOn = [];            // CC habilitados, ordenados por código: los únicos que se pueden asignar
let rows = [];
const P = () => S.periodos.find(p => p.id === S.actual) || null;

/* ================= Carga desde Supabase ================= */
async function loadMaestros() {
  const [per, cc, zs, bases, asig] = await Promise.all([
    sb.from("seg_periodos").select("*").order("creado_en", { ascending: false }),
    fetchAll(() => sb.from("seg_centros_costo").select("codigo,descripcion,habilitado").order("codigo")),
    fetchAll(() => sb.from("seg_zonas").select("patente,zona").order("patente")),
    sb.from("seg_bases").select("*"),
    fetchAll(() => sb.from("seg_asignaciones").select("patente").order("patente"))
  ]);
  if (per.error) throw per.error;
  if (bases.error) throw bases.error;
  S.periodos = per.data;
  S.ccBase = cc;
  S.zonas = Object.fromEntries(zs.map(z => [z.patente, z.zona]));
  S.bases = Object.fromEntries(bases.data.map(b => [b.tipo, b]));
  S.asignadas = new Set(asig.map(a => a.patente));
  await loadPdfInfo();
  const saved = localStorage.getItem("seg-periodo");
  if (!S.periodos.some(p => p.id === S.actual)) S.actual = S.periodos.some(p => p.id === saved) ? saved : (S.periodos[0]?.id || null);
  buildCC();
}
async function loadPeriodo() {
  rows = [];
  const p = P(); if (!p) return;
  const data = await fetchAll(() => sb.from("seg_v_prorrateo").select("*").eq("periodo_id", p.id).order("certificado"));
  rows = data.map(r => ({
    cert: r.certificado, patente: r.patente, marca: r.marca, anio: r.anio || "", zona: r.zona || "",
    valor: +r.valor, premio: +r.premio, base: +r.premio, pct: +r.porcentaje,
    neto: +r.neto, iva: +r.iva, perc: +r.percepcion, otros: +r.otros_impuestos, total: +r.total,
    cc: r.cc || "", nuevo: r.nueva
  }));
}
async function reload(opts = {}) {
  try {
    if (!opts.soloPeriodo) await loadMaestros();
    await loadPeriodo();
    renderAll();
  } catch (e) { toast("No se pudieron leer los datos: " + errMsg(e)); }
}

/* ================= CC ================= */
function buildCC() {
  ccMap = new Map(S.ccBase.map(c => [c.codigo, { d: c.descripcion, on: c.habilitado }]));
  ccOn = S.ccBase.filter(c => c.habilitado)
    .map(c => ({ code: c.codigo, d: c.descripcion || "", key: normH(c.codigo + " " + (c.descripcion || "")) }))
    .sort((a, b) => a.code.localeCompare(b.code, "es", { numeric: true }));
}
const ccHabilitado = code => !!(code && ccMap.get(code)?.on);
function codeFrom(val) {
  val = String(val || "").trim(); if (!val) return "";
  const code = val.split(/\s+·\s+|\s{2,}/)[0].trim();
  return ccMap.has(code) ? code : (ccMap.has(val) ? val : code);
}
function ccStatus(code) {
  if (!code) return { cls: "", txt: "Sin asignar" };
  const c = ccMap.get(code);
  if (!c) return { cls: "bad", txt: "No existe en la base de CC" };
  if (!c.on) return { cls: "ok off", txt: c.d + " · inhabilitado" };
  return { cls: "ok", txt: c.d };
}
function groupByCC() {
  const g = new Map();
  for (const r of rows) {
    const k = r.cc || "";
    if (!g.has(k)) g.set(k, { cc: k, n: 0, neto: 0, iva: 0, perc: 0, otros: 0, total: 0, pats: [] });
    const x = g.get(k); x.n++; x.pats.push(r); COMP.concat("total").forEach(c => x[c] += r[c]);
  }
  for (const x of g.values()) COMP.concat("total").forEach(c => x[c] = Math.round(x[c] * 100) / 100);
  return [...g.values()].sort((a, b) => (b.cc === "" ? 1 : 0) - (a.cc === "" ? 1 : 0) || b.total - a.total);
}

/* ================= Render ================= */
function renderAll() {
  buildCC();
  const p = P();
  const sel = $("per");
  sel.hidden = S.periodos.length < 2; $("newper").hidden = !p; $("exp").hidden = !p; $("tabs").hidden = !p;
  sel.innerHTML = S.periodos.map(x => `<option value="${esc(x.id)}" ${x.id === S.actual ? "selected" : ""}>${esc(isoToAr(x.vigencia_desde))} al ${esc(isoToAr(x.vigencia_hasta))} · Fact. ${esc(x.id)}</option>`).join("");
  if (!p) { $("main").hidden = true; $("invoice").hidden = true; showUpload(true); $("hsub").textContent = "Cargá la póliza, la factura y la base de CC para ver cuánto le toca a cada centro de costo."; return; }
  if ($("upload").dataset.forced !== "1") $("upload").hidden = true;
  $("main").hidden = false;
  $("hsub").textContent = `Póliza ${p.poliza || ""} · ${p.tomador || ""} · Vigencia ${isoToAr(p.vigencia_desde)} al ${isoToAr(p.vigencia_hasta)} · ${rows.length} vehículos`;
  renderInvoice(p); renderProgress(); renderDash(); renderCC(); renderPat(); setView(S.view);
}
function renderInvoice(p) {
  $("i-num").textContent = "Factura " + p.id;
  renderPdfBtn(p);
  $("i-dl").innerHTML = `<dt>Fecha</dt><dd>${esc(isoToAr(p.factura_fecha))}</dd><dt>Detalle</dt><dd>${esc(p.factura_detalle || "")}</dd>`;
  const T = +p.neto + +p.iva + +p.percepcion + +p.otros_impuestos;
  const cuadra = Math.abs(rows.reduce((a, r) => a + r.total, 0) - T) < 0.005;
  $("i-amts").innerHTML = [
    ["Neto (subtotal)", p.neto, "Prima + recargo financiero"], ["IVA", p.iva, "21% + IVA s/recargo financ."],
    ["Percepción IVA", p.percepcion, "IVA 3%"], ["Otros impuestos", p.otros_impuestos, "Imp. y tasas, sellos, C.S."]
  ].map(([l, v, h]) => `<div class="amt"><div class="l">${l}</div><div class="v">$ ${fmt(+v)}</div><div class="hint">${h}</div></div>`).join("")
    + `<div class="amt total"><div class="l">Total factura</div><div class="v">$ ${fmt(T)}</div><div class="hint">${cuadra ? "Prorrateo cuadra al centavo" : "Revisar: el prorrateo no cuadra"}</div></div>`;
}
function renderProgress() {
  const done = rows.filter(r => r.cc && ccMap.has(r.cc)).length, pend = rows.length - done;
  $("pg-txt").textContent = `${done} de ${rows.length}`;
  $("pg-bar").style.width = (rows.length ? done / rows.length * 100 : 0) + "%";
  $("t-cnt").hidden = !pend; $("t-cnt").innerHTML = `${pend}<span class="t-long"> pendientes</span>`; $("t-cnt").title = pend + " patentes sin centro de costo";
  $("pg-hint").textContent = !rows.length ? "" : pend ? `${pend} sin CC · ver cuáles` : "Todas con centro de costo";
  $("asig").classList.toggle("pend", pend > 0);
  $("asig").disabled = !pend;
}
function renderCC() {
  const g = groupByCC(); const T = rows.reduce((a, r) => a + r.total, 0);
  const max = Math.max(...g.map(x => x.total), 1);
  $("ccb").innerHTML = g.map(x => {
    const c = ccMap.get(x.cc), key = x.cc || "__none";
    const cls = x.cc === "" ? "none" : !c ? "bad" : "";
    const desc = x.cc === "" ? "Patentes sin centro de costo asignado" : (c ? c.d + (c.on ? "" : " (inhabilitado)") : "No existe en la base de CC");
    const zs = x.cc ? [...new Set(x.pats.map(r => r.zona).filter(Boolean))] : [];
    const open = S.open.has(key);
    let html = `<tr class="ccrow ${cls} ${open ? "open" : ""}" data-k="${esc(key)}" tabindex="0" aria-expanded="${open}">
      <td class="code">${x.cc ? esc(x.cc) : "Sin CC"}</td><td class="desc">${esc(desc)}${zs.length ? `<small class="zl">${esc(zs.join(", "))}</small>` : ""}</td>
      <td class="sh"><div class="share"><div class="b"><i style="width:${(x.total / max * 100).toFixed(1)}%"></i></div><span>${fmtPct(T ? x.total / T : 0)}</span></div></td>
      <td class="num n" data-l="Patentes">${x.n}</td><td class="num" data-l="Neto">${fmt(x.neto)}</td><td class="num" data-l="IVA">${fmt(x.iva)}</td><td class="num" data-l="Perc. IVA">${fmt(x.perc)}</td><td class="num" data-l="Otros imp.">${fmt(x.otros)}</td><td class="num big">$ ${fmt(x.total)}</td></tr>`;
    if (open) {
      html += `<tr class="detail"><td colspan="9"><div class="chips">${x.pats.slice().sort((a, b) => b.total - a.total).map(r => `<span class="chip"><span class="plate sm"><span>${esc(r.patente)}</span></span><em>${esc(r.marca)}</em>${r.zona ? `<span class="zona">${esc(r.zona)}</span>` : ""}<b>$ ${fmt(r.total)}</b></span>`).join("")}</div>
      ${x.cc === "" ? `<div style="margin-top:10px"><button class="btn primary" type="button" data-go="none">Asignar CC a estas patentes</button></div>` : ""}</td></tr>`;
    }
    return html;
  }).join("");
  const s = k => rows.reduce((a, r) => a + r[k], 0);
  $("ccf").innerHTML = `<tr><td class="code">Total</td><td class="desc">${g.filter(x => x.cc).length} centros de costo</td><td class="sh"></td><td class="num n" data-l="Patentes">${rows.length}</td><td class="num" data-l="Neto">${fmt(s("neto"))}</td><td class="num" data-l="IVA">${fmt(s("iva"))}</td><td class="num" data-l="Perc. IVA">${fmt(s("perc"))}</td><td class="num" data-l="Otros imp.">${fmt(s("otros"))}</td><td class="num big">$ ${fmt(s("total"))}</td></tr>`;
}
/* Filtros encadenados: cada lista muestra solo lo que existe con los demás filtros aplicados */
const FILS = [["all", "Todas las patentes"], ["none", "Sin centro de costo"], ["done", "Con centro de costo"], ["bad", "Con CC inválido o inhabilitado"], ["new", "Nuevas en este período"]];
function matchQ(r) {
  const q = S.q.trim().toLowerCase(); if (!q) return true;
  const cd = ccMap.get(r.cc);
  return (r.patente + " " + r.marca + " " + (r.zona || "") + " " + (r.cc || "") + " " + (cd ? cd.d : "")).toLowerCase().includes(q);
}
function matchFil(r, f) {
  if (f === "none") return !r.cc;
  if (f === "done") return !!r.cc;
  if (f === "bad") { const st = ccStatus(r.cc).cls; return st === "bad" || st.includes("off"); }
  if (f === "new") return !!r.nuevo;
  return true;
}
function matchZ(r, z) {
  if (!z) return true;
  return z === "__none" ? !r.zona : r.zona === z;
}
function visible() {
  const v = rows.filter(r => matchQ(r) && matchFil(r, S.fil) && matchZ(r, S.zf));
  const { k, dir } = S.sort;
  return v.sort((a, b) => { const x = a[k] ?? "", y = b[k] ?? ""; return (typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "es")) * dir; });
}
function fillZonas() {
  // Zonas de los vehículos que cumplen la búsqueda y el filtro de CC
  const base = rows.filter(r => matchQ(r) && matchFil(r, S.fil));
  const cnt = new Map(); let sinZ = 0;
  base.forEach(r => { if (r.zona) cnt.set(r.zona, (cnt.get(r.zona) || 0) + 1); else sinZ++; });
  const zs = [...cnt.keys()].sort((a, b) => a.localeCompare(b, "es"));
  if (S.zf && S.zf !== "__none" && !cnt.has(S.zf)) S.zf = "";
  if (S.zf === "__none" && !sinZ) S.zf = "";
  $("zf").innerHTML = (base.length ? `<option value="">Todas las zonas (${base.length})</option>` : `<option value="">Sin zonas para este filtro</option>`)
    + zs.map(z => `<option value="${esc(z)}">${esc(z)} (${cnt.get(z)})</option>`).join("")
    + (sinZ ? `<option value="__none">Sin zona (${sinZ})</option>` : "");
  $("zf").value = S.zf;
  $("zf").disabled = !base.length;
}
function fillFil() {
  // Estados de CC de los vehículos que cumplen la búsqueda y la zona
  const base = rows.filter(r => matchQ(r) && matchZ(r, S.zf));
  $("fil").innerHTML = FILS.map(([v, l]) => {
    const n = base.filter(r => matchFil(r, v)).length;
    const dis = n === 0 && v !== "all" && v !== S.fil;
    return `<option value="${v}" ${dis ? "disabled" : ""}>${l} (${n})</option>`;
  }).join("");
  $("fil").value = S.fil;
  $("fclear").hidden = !(S.fil !== "all" || S.zf || S.q.trim());
}
function renderPat() {
  fillZonas(); fillFil();
  const v = visible();
  $("tb").innerHTML = !v.length ? `<tr><td colspan="13" class="empty">Ninguna patente coincide con la búsqueda o el filtro.</td></tr>` :
    v.map(r => {
      const st = ccStatus(r.cc); return `<tr data-p="${esc(r.patente)}" class="${S.sel.has(r.patente) ? "sel" : ""}">
      <td class="chk"><input type="checkbox" class="rs" ${S.sel.has(r.patente) ? "checked" : ""} aria-label="Seleccionar ${esc(r.patente)}"></td>
      <td class="cert">${r.cert ?? ""}</td>
      <td class="pat"><span class="plate"><span>${esc(r.patente)}</span></span>${r.nuevo ? '<span class="new">Nueva</span>' : ""}</td>
      <td class="marca">${esc(r.marca)}<small>${esc(r.anio)}</small></td>
      <td class="zn">${r.zona ? `<span class="zona">${esc(r.zona)}</span>` : `<span class="zona none">Sin zona</span>`}</td>
      <td class="ccc"><div class="cc ${st.cls}"><button type="button" class="ccbtn${r.cc ? "" : " empty"}" aria-haspopup="dialog" aria-expanded="false" aria-label="Centro de costo de ${esc(r.patente)}: ${esc(r.cc ? r.cc + ", " + st.txt : "sin asignar")}"><span class="v">${r.cc ? esc(r.cc) : "Elegir CC"}</span><span class="arr" aria-hidden="true"></span></button><small title="${esc(st.txt)}">${esc(st.txt)}</small></div></td>
      <td class="num" data-l="Premio">${fmt(r.base)}</td><td class="num" data-l="%">${fmtPct(r.pct)}</td>
      <td class="num" data-l="Neto">${fmt(r.neto)}</td><td class="num" data-l="IVA">${fmt(r.iva)}</td><td class="num" data-l="Perc. IVA">${fmt(r.perc)}</td><td class="num" data-l="Otros imp.">${fmt(r.otros)}</td>
      <td class="num tot" data-l="Total">${fmt(r.total)}</td></tr>`;
    }).join("");
  const s = k => v.reduce((a, r) => a + r[k], 0);
  $("tf").innerHTML = !v.length ? "" : `<tr><td class="chk"></td><td class="lbl" colspan="5">${v.length !== rows.length ? `Subtotal ${v.length} patentes visibles` : `Total ${rows.length} patentes`}</td><td class="num" data-l="Premio">${fmt(s("base"))}</td><td class="num" data-l="%">${fmtPct(s("pct"))}</td><td class="num" data-l="Neto">${fmt(s("neto"))}</td><td class="num" data-l="IVA">${fmt(s("iva"))}</td><td class="num" data-l="Perc. IVA">${fmt(s("perc"))}</td><td class="num" data-l="Otros imp.">${fmt(s("otros"))}</td><td class="num tot" data-l="Total">${fmt(s("total"))}</td></tr>`;
  document.querySelectorAll("#patview th[data-sort]").forEach(th => {
    const on = th.dataset.sort === S.sort.k;
    th.classList.toggle("asc", on && S.sort.dir === 1); th.classList.toggle("desc", on && S.sort.dir === -1);
    th.setAttribute("aria-sort", on ? (S.sort.dir === 1 ? "ascending" : "descending") : "none");
  });
  const sv = S.sort.k + ":" + S.sort.dir, ss = $("sortsel");
  if ([...ss.options].some(o => o.value === sv)) ss.value = sv;
  syncBulk();
}
function setView(v) {
  S.view = v;
  $("dashview").hidden = v !== "dash"; $("tblpanel").hidden = v === "dash";
  $("ccview").hidden = v !== "cc"; $("patview").hidden = v !== "pat";
  // En el Dashboard los importes de la factura ya están en los indicadores
  $("invoice").hidden = !P() || v === "dash";
  ["dash", "cc", "pat"].forEach(x => $("t-" + x).setAttribute("aria-selected", v === x));
}
function refreshViews() { const y = scrollY; renderProgress(); renderDash(); renderCC(); renderPat(); scrollTo(0, y); }

/* ================= Asignación ================= */
async function assign(pats, code) {
  const p = P(); if (!p || !pats.length) return false;
  if (code && !ccHabilitado(code)) { toast(ccMap.has(code) ? `El CC ${code} está inhabilitado; elegí uno habilitado.` : `El CC ${code} no está en la base.`); return false; }
  const prev = new Map(rows.filter(r => pats.includes(r.patente)).map(r => [r.patente, { cc: r.cc, nuevo: r.nuevo }]));
  rows.forEach(r => { if (prev.has(r.patente)) { r.cc = code; r.nuevo = false; } });
  refreshViews();
  try {
    await rpc("seg_asignar_cc", { p_periodo: p.id, p_patentes: pats, p_cc: code || null });
    pats.forEach(pt => code ? S.asignadas.add(pt) : S.asignadas.delete(pt));
    return true;
  } catch (e) {
    rows.forEach(r => { if (prev.has(r.patente)) Object.assign(r, prev.get(r.patente)); });
    refreshViews(); toast("No se guardó el CC: " + errMsg(e)); return false;
  }
}
$("tb").addEventListener("change", e => {
  const tr = e.target.closest("tr"); if (!tr) return; const pt = tr.dataset.p;
  if (e.target.classList.contains("rs")) { e.target.checked ? S.sel.add(pt) : S.sel.delete(pt); tr.classList.toggle("sel", e.target.checked); syncBulk(); }
});
const rowBtn = pt => $("tb").querySelector(`tr[data-p="${CSS.escape(pt)}"] .ccbtn`);
function pickRow(btn) {
  const pt = btn.closest("tr").dataset.p, r = rows.find(x => x.patente === pt); if (!r) return;
  openPicker(btn, {
    title: "Centro de costo de " + pt, cur: r.cc, allowNone: true,
    onPick: async code => { if (code !== r.cc) await assign([pt], code); },
    refocus: () => rowBtn(pt)
  });
}
$("tb").addEventListener("click", e => { const b = e.target.closest(".ccbtn"); if (b) pickRow(b); });
$("tb").addEventListener("keydown", e => { const b = e.target.closest(".ccbtn"); if (b && (e.key === "ArrowDown" || e.key === "F4")) { e.preventDefault(); pickRow(b); } });
function syncBulk() {
  const n = S.sel.size; $("bulk").classList.toggle("on", n > 0); document.body.classList.toggle("bulk-on", n > 0); setBulkH(); $("bulk-n").textContent = n === 1 ? "1 seleccionada" : `${n} seleccionadas`;
  const vis = [...$("tb").querySelectorAll("tr[data-p]")].map(t => t.dataset.p), all = $("all");
  all.checked = vis.length > 0 && vis.every(p => S.sel.has(p)); all.indeterminate = !all.checked && vis.some(p => S.sel.has(p));
}
$("all").addEventListener("change", e => { [...$("tb").querySelectorAll("tr[data-p]")].forEach(t => e.target.checked ? S.sel.add(t.dataset.p) : S.sel.delete(t.dataset.p)); renderPat(); });
function renderBulkCC() {
  const b = $("bulk-cc"), c = ccMap.get(S.bulkCC);
  b.classList.toggle("empty", !S.bulkCC);
  b.querySelector(".v").innerHTML = S.bulkCC ? `${esc(S.bulkCC)}<em>${esc(c ? c.d : "")}</em>` : "Centro de costo para todas";
  b.setAttribute("aria-label", "Centro de costo para las seleccionadas: " + (S.bulkCC ? S.bulkCC + (c ? ", " + c.d : "") : "sin elegir"));
}
function pickBulk() {
  openPicker($("bulk-cc"), {
    title: S.sel.size === 1 ? "Centro de costo para 1 patente" : `Centro de costo para ${S.sel.size} patentes`, cur: S.bulkCC, allowNone: false,
    onPick: code => { S.bulkCC = code; renderBulkCC(); },
    refocus: () => $("bulk-ok")
  });
}
$("bulk-cc").addEventListener("click", pickBulk);
$("bulk-cc").addEventListener("keydown", e => { if (e.key === "ArrowDown" || e.key === "F4") { e.preventDefault(); pickBulk(); } });
$("bulk-ok").addEventListener("click", async () => {
  const code = S.bulkCC;
  if (!code) { toast("Elegí un centro de costo para asignar."); return pickBulk(); }
  const pats = [...S.sel], btn = $("bulk-ok"); busy(btn, true);
  if (await assign(pats, code)) { toast(`CC ${code} asignado a ${pats.length} patentes.`); S.sel.clear(); S.bulkCC = ""; renderBulkCC(); refreshViews(); }
  busy(btn, false);
});
$("bulk-clear").addEventListener("click", async () => { const pats = [...S.sel]; if (await assign(pats, "")) { S.sel.clear(); refreshViews(); } });
$("bulk-x").addEventListener("click", () => { S.sel.clear(); S.bulkCC = ""; renderBulkCC(); renderPat(); });
document.querySelectorAll("#patview th[data-sort]").forEach(th => {
  th.tabIndex = 0;
  const go = () => { const k = th.dataset.sort; S.sort = { k, dir: S.sort.k === k ? -S.sort.dir : 1 }; renderPat(); };
  th.addEventListener("click", go);
  th.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); } });
});
$("sortsel").addEventListener("change", e => { const [k, d] = e.target.value.split(":"); S.sort = { k, dir: +d }; renderPat(); });
$("q").addEventListener("input", e => { S.q = e.target.value; renderPat(); });
$("fil").addEventListener("change", e => { S.fil = e.target.value; S.sel.clear(); renderPat(); });
$("fclear").addEventListener("click", () => { S.fil = "all"; S.zf = ""; S.q = ""; $("q").value = ""; S.sel.clear(); renderPat(); });
$("zf").addEventListener("change", e => { S.zf = e.target.value; S.sel.clear(); renderPat(); });
$("t-dash").addEventListener("click", () => setView("dash"));
$("t-cc").addEventListener("click", () => setView("cc"));
$("t-pat").addEventListener("click", () => setView("pat"));
function toggleCC(row) { const k = row.dataset.k; S.open.has(k) ? S.open.delete(k) : S.open.add(k); renderCC(); }
$("ccb").addEventListener("click", e => {
  if (e.target.closest("[data-go]")) { S.fil = "none"; S.zf = ""; S.sel.clear(); setView("pat"); renderPat(); return; }
  const row = e.target.closest(".ccrow"); if (row) toggleCC(row);
});
$("ccb").addEventListener("keydown", e => { const row = e.target.closest(".ccrow"); if (row && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); toggleCC(row); } });
$("asig").addEventListener("click", () => { S.fil = "none"; S.zf = ""; S.q = ""; $("q").value = ""; S.sel.clear(); setView("pat"); renderPat(); });
$("per").addEventListener("change", async e => { S.actual = e.target.value; localStorage.setItem("seg-periodo", S.actual); S.sel.clear(); S.open.clear(); await reload({ soloPeriodo: true }); });

/* ================= Dashboard =================
   Todo se calcula con los datos guardados: las patentes del período visible
   (seg_v_prorrateo), la factura y el historial de seg_periodos. */
const MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
const pct1 = v => (v * 100).toLocaleString("es-AR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "%";
const fmt0 = n => "$ " + Math.round(n || 0).toLocaleString("es-AR");
const fmtM = n => "$ " + (n / 1e6).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " M";
const fmtAxis = n => n >= 1e6 ? (n / 1e6).toLocaleString("es-AR", { maximumFractionDigits: 1 }) + " M" : n >= 1e3 ? (n / 1e3).toLocaleString("es-AR", { maximumFractionDigits: 0 }) + " k" : String(n);
const fecha = d => { if (!d) return null; const [y, m, dd] = String(d).slice(0, 10).split("-").map(Number); return new Date(y, m - 1, dd); };
const fechaAr = d => { const x = d instanceof Date ? d : fecha(d); return x ? x.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric" }) : ""; };
const totFact = p => +p.neto + +p.iva + +p.percepcion + +p.otros_impuestos;
const cap = s => s.charAt(0).toUpperCase() + s.slice(1);
function perLabel(p) {
  const a = fecha(p.vigencia_desde), b = fecha(p.vigencia_hasta);
  if (!a || !b) return "Factura " + p.id;
  return a.getFullYear() === b.getFullYear()
    ? `${cap(MESES[a.getMonth()])} – ${MESES[b.getMonth()]} ${b.getFullYear()}`
    : `${cap(MESES[a.getMonth()])} ${a.getFullYear()} – ${MESES[b.getMonth()]} ${b.getFullYear()}`;
}
const ccLabel = cc => { const c = ccMap.get(cc); return c && c.d ? `CC ${cc} · ${c.d}` : `CC ${cc}`; };
function niceMax(v) {
  if (v <= 0) return { max: 1, step: 1 };
  const raw = v / 4, mag = 10 ** Math.floor(Math.log10(raw)), f = raw / mag;
  const step = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * mag;
  return { max: Math.ceil(v / step) * step, step };
}
const ICON = {
  money: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 3h6l-1.6 3.2C17.3 7.6 20 11.6 20 15.5 20 19.4 16.8 21 12 21s-8-1.6-8-5.5c0-3.9 2.7-7.9 6.6-9.3z" fill="currentColor"/><text x="12" y="17.6" text-anchor="middle" font-size="9.5" font-weight="700" font-family="Arial,sans-serif" fill="var(--kc)">$</text></svg>',
  car: '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M4 13.2 5.8 8a2.4 2.4 0 0 1 2.3-1.6h7.8A2.4 2.4 0 0 1 18.2 8L20 13.2V19a1 1 0 0 1-1 1h-1.5a1 1 0 0 1-1-1v-1.2h-9V19a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1zm2.3-.4h11.4L16.4 8.8a.9.9 0 0 0-.9-.6h-7a.9.9 0 0 0-.9.6zM7.3 16a1.3 1.3 0 1 0 0-2.6 1.3 1.3 0 0 0 0 2.6zm9.4 0a1.3 1.3 0 1 0 0-2.6 1.3 1.3 0 0 0 0 2.6z"/></svg>',
  cc: '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M5 3h9a1 1 0 0 1 1 1v5h4a1 1 0 0 1 1 1v11H4V4a1 1 0 0 1 1-1zm2 3v2h2V6zm4 0v2h2V6zm-4 4v2h2v-2zm4 0v2h2v-2zm-4 4v2h2v-2zm4 0v2h2v-2zm5-3v2h2v-2zm0 4v2h2v-2z"/></svg>',
  check: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9" fill="currentColor"/><path d="m7.6 12.3 3 3 5.8-6.1" fill="none" stroke="var(--kc)" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  warn: '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 2.5 22.5 20.5h-21z"/><path d="M12 9v5.2" stroke="var(--kc)" stroke-width="2.3" stroke-linecap="round"/><circle cx="12" cy="17.3" r="1.3" fill="var(--kc)"/></svg>',
  cal: '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M7 2h2v2h6V2h2v2h2a2 2 0 0 1 2 2v13a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2zM5 9v10h14V9zm2 2h2v2H7zm4 0h2v2h-2zm4 0h2v2h-2zm-8 4h2v2H7zm4 0h2v2h-2z"/></svg>',
  eye: '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 4.5c4.2 0 7.3 3.6 8.3 5.5-1 1.9-4.1 5.5-8.3 5.5S2.7 11.9 1.7 10C2.7 8.1 5.8 4.5 10 4.5zm0 2.5a3 3 0 1 0 0 6 3 3 0 0 0 0-6z" fill="currentColor"/></svg>',
  chev: '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="m7.5 4.5 5.5 5.5-5.5 5.5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  a_warn: '<svg viewBox="0 0 20 20" aria-hidden="true"><path fill="currentColor" d="M10 1.8 19 17.6H1z"/><path d="M10 7.3v4.6" stroke="var(--ac-bg)" stroke-width="2" stroke-linecap="round"/><circle cx="10" cy="14.4" r="1.1" fill="var(--ac-bg)"/></svg>',
  a_bad: '<svg viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="8.5" fill="currentColor"/><path d="m6.8 6.8 6.4 6.4m0-6.4-6.4 6.4" stroke="var(--ac-bg)" stroke-width="2" stroke-linecap="round"/></svg>',
  a_ok: '<svg viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="8.5" fill="currentColor"/><path d="m6.2 10.3 2.5 2.5 5-5.3" fill="none" stroke="var(--ac-bg)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  a_info: '<svg viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="8.5" fill="currentColor"/><path d="M10 9v5" stroke="var(--ac-bg)" stroke-width="2" stroke-linecap="round"/><circle cx="10" cy="6.2" r="1.15" fill="var(--ac-bg)"/></svg>'
};

// 12 colores de la paleta; si hay más CC se repiten en tonos más claros y más oscuros
function serieColor(i) {
  const base = `var(--s${(i % 12) + 1})`, vuelta = Math.floor(i / 12);
  if (!vuelta) return base;
  return vuelta % 2 ? `color-mix(in srgb, ${base} 55%, var(--surface))` : `color-mix(in srgb, ${base} 65%, var(--ink))`;
}
function dashData() {
  const p = P(), n = rows.length, F = totFact(p), T = rows.reduce((a, r) => a + r.total, 0);
  const sin = rows.filter(r => !r.cc).length;
  const val = rows.filter(r => ccHabilitado(r.cc)).length;
  const inc = n - sin - val;
  // Todos los centros de costo, ordenados por monto (se usan para cargar el prorrateo en Bejerman)
  const g = groupByCC(), con = g.filter(x => x.cc).sort((a, b) => b.total - a.total);
  const items = con.map((x, i) => ({ key: x.cc, label: ccLabel(x.cc), total: x.total, color: serieColor(i), bad: !ccHabilitado(x.cc) }));
  const none = g.find(x => !x.cc);
  if (none) items.push({ key: "__none", label: "Sin centro de costo", total: none.total, color: "var(--warn)", none: true });
  return { p, n, F, T, sin, val, inc, items, ccUsados: con.length, nuevas: rows.filter(r => r.nuevo).length, sinZona: rows.filter(r => !r.zona).length };
}
function donut(segs, size, stroke, center) {
  const r = (size - stroke) / 2, C = 2 * Math.PI * r, tot = segs.reduce((a, s) => a + s.v, 0);
  let off = 0;
  const arcs = !tot ? `<circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="var(--barbg)" stroke-width="${stroke}"/>`
    : segs.filter(s => s.v > 0).map(s => {
      const len = s.v / tot * C, gap = segs.filter(x => x.v > 0).length > 1 ? Math.min(1.5, len / 3) : 0;
      const el = `<circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" style="stroke:${s.color}" stroke-width="${stroke}" stroke-dasharray="${(len - gap).toFixed(2)} ${(C - len + gap).toFixed(2)}" stroke-dashoffset="${(-off).toFixed(2)}" transform="rotate(-90 ${size / 2} ${size / 2})"><title>${esc(s.title || "")}</title></circle>`;
      off += len; return el;
    }).join("");
  return `<div class="donut" style="--ds:${size}px"><svg viewBox="0 0 ${size} ${size}" role="img" aria-label="${esc(center.aria)}">${arcs}</svg><div class="dc"><b>${center.v}</b><span>${center.l}</span></div></div>`;
}
function renderDash() {
  const p = P(); if (!p) return;
  const d = dashData();
  $("d-per").textContent = `${perLabel(p)} · Factura ${p.id}`;

  // Vigencia: días que faltan (o si ya venció / todavía no empezó)
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const desde = fecha(p.vigencia_desde), hasta = fecha(p.vigencia_hasta), DIA = 864e5;
  let vig = { v: "—", l: "Vigencia", s: "Sin fechas en la póliza" };
  if (desde && hasta) {
    const total = Math.round((hasta - desde) / DIA) + 1;
    if (hoy < desde) vig = { v: `${Math.round((desde - hoy) / DIA)} días`, l: "Para el inicio", s: `Vigencia ${fechaAr(desde)} al ${fechaAr(hasta)}` };
    else if (hoy > hasta) vig = { v: "Vencida", l: "Vigencia", s: `Venció el ${fechaAr(hasta)} · ${total} días` };
    else { const f = Math.round((hasta - hoy) / DIA); vig = { v: `${f} ${f === 1 ? "día" : "días"}`, l: "Restan de vigencia", s: `Vence el ${fechaAr(hasta)} · ${total} días en total` }; }
  }
  const pend = d.sin + d.inc;
  const kpi = (cls, icon, v, l, s, go) => `<${go ? `button type="button" data-go="${go}"` : "div"} class="kpi ${cls}"><span class="ki">${ICON[icon]}</span><div><div class="kv" title="${String(v).replace(/<[^>]*>/g, "")}">${v}</div><div class="kl">${l}</div><div class="ks">${s}</div></div>${go ? `<span class="kgo" aria-hidden="true">${ICON.chev}</span>` : ""}</${go ? "button" : "div"}>`;
  $("d-kpis").innerHTML =
    kpi("k-money", "money", fmt0(d.F), "Costo total", "Total de la factura") +
    kpi("k-car", "car", fmtInt(d.n), "Vehículos asegurados", d.nuevas ? `${d.nuevas} ${d.nuevas === 1 ? "nuevo" : "nuevos"} en este período` : `Póliza ${esc(p.poliza || "")}`) +
    kpi("k-cc", "cc", fmtInt(d.ccUsados), "Centros de costo", `con vehículos · ${fmtInt(ccOn.length)} habilitados`, "cc") +
    kpi("k-check", "check", pct1(d.n ? d.val / d.n : 0), "Patentes asignadas", `${fmtInt(d.val)} de ${fmtInt(d.n)} con CC habilitado`) +
    kpi("k-warn" + (pend ? " pend" : ""), "warn", fmtInt(pend), "Sin asignación", pend ? [d.sin && `${d.sin} sin CC`, d.inc && `${d.inc} con CC no habilitado`].filter(Boolean).join(" · ") : "Todas tienen CC", pend ? (d.sin ? "none" : "bad") : "") +
    kpi("k-cal", "cal", vig.v, vig.l, vig.s);

  // Barras: costo por centro de costo
  const max = Math.max(...d.items.map(x => x.total), 0), ax = niceMax(max);
  const ticks = []; for (let t = 0; t <= ax.max + 1e-9; t += ax.step) ticks.push(t);
  $("d-bars").innerHTML = !d.items.length ? `<p class="dempty">El período no tiene patentes.</p>` :
    `<div class="bars">${d.items.map(x => `<button type="button" class="brow${x.none ? " none" : ""}${x.bad ? " bad" : ""}" data-cc="${esc(x.key)}" title="Ver ${esc(x.label)} en Totales por CC">
        <span class="bl">${esc(x.label)}${x.bad ? `<small>${ccMap.has(x.key) ? "inhabilitado" : "no existe en la base"}</small>` : ""}</span>
        <span class="bt"><i style="--w:${(x.total / ax.max).toFixed(4)};background:${x.color}"></i><em>$ ${fmt(x.total)}</em></span>
        <span class="bp">${pct1(d.T ? x.total / d.T : 0)}</span></button>`).join("")}
      <div class="baxis" aria-hidden="true"><span></span><span class="bt">${ticks.map(t => `<b style="--w:${(t / ax.max).toFixed(4)}">${t ? fmtAxis(t) : "0"}</b>`).join("")}</span><span></span></div></div>`;

  // Anillo: participación
  $("d-donut").innerHTML = !d.items.length ? `<p class="dempty">El período no tiene patentes.</p>` :
    `<div class="dsplit">${donut(d.items.map(x => ({ v: x.total, color: x.color, title: `${x.label}: ${fmt0(x.total)}` })), 190, 40, { v: fmtM(d.T), l: "Total", aria: `Participación por centro de costo, total ${fmt0(d.T)}` })}
      <ul class="legend full">${d.items.map(x => `<li><i style="background:${x.color}"></i><span>${esc(x.label)}</span><em>$ ${fmt(x.total)}</em><b>${pct1(d.T ? x.total / d.T : 0)}</b></li>`).join("")}
        <li class="lt"><i></i><span>Total · ${d.items.filter(x => !x.none).length} centros de costo${d.items.some(x => x.none) ? " y sin CC" : ""}</span><em>$ ${fmt(d.T)}</em><b>100%</b></li></ul></div>`;

  // Estado de patentes
  const pct = d.n ? d.val / d.n : 0;
  $("d-estado").innerHTML = `<div class="dsplit sm">${donut([
      { v: d.val, color: "var(--ok)", title: `Asignadas: ${d.val}` },
      { v: d.sin, color: "var(--warn)", title: `Sin asignación: ${d.sin}` },
      { v: d.inc, color: "var(--bad)", title: `Con CC no habilitado: ${d.inc}` }], 150, 22,
      { v: pct1(pct), l: "Asignadas", aria: `${fmtPct(pct)} de las patentes asignadas` })}
    <ul class="legend counts">
      <li><i style="background:var(--ok)"></i><span>Asignadas</span><b>${fmtInt(d.val)}</b></li>
      <li><i style="background:var(--warn)"></i><span>Sin asignación</span><b>${fmtInt(d.sin)}</b></li>
      <li><i style="background:var(--bad)"></i><span>Con CC no habilitado</span><b>${fmtInt(d.inc)}</b></li>
    </ul></div>
    ${pend ? `<button type="button" class="dnote warn" data-go="${d.sin ? "none" : "bad"}">${ICON.a_warn}<span>${pend === 1 ? "1 patente necesita" : `${pend} patentes necesitan`} un centro de costo habilitado</span>${ICON.chev}</button>`
      : `<div class="dnote ok">${ICON.a_ok}<span>Todas las patentes tienen un centro de costo habilitado</span></div>`}`;

  // Últimos períodos, del más reciente al más antiguo, con la variación contra el anterior
  const pers = S.periodos.slice().sort((a, b) => String(b.vigencia_desde || b.creado_en).localeCompare(String(a.vigencia_desde || a.creado_en)));
  const shown = S.dashAll ? pers : pers.slice(0, 3);
  $("d-pers").innerHTML = `<div class="ptable"><table><thead><tr><th>Período</th><th class="num">Costo total</th><th class="num">Var. vs. anterior</th><th class="pa"><span class="vh">Ver</span></th></tr></thead><tbody>
    ${shown.map(x => {
      const i = pers.indexOf(x), prev = pers[i + 1], t = totFact(x), v = prev ? t / totFact(prev) - 1 : null, cur = x.id === S.actual;
      const vt = v == null ? `<span class="var">—</span>` : `<span class="var ${v > 0.0005 ? "up" : v < -0.0005 ? "down" : ""}">${v > 0 ? "+" : ""}${(v * 100).toLocaleString("es-AR", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%</span>`;
      return `<tr class="${cur ? "cur" : ""}"><td data-l="Período"><b>${esc(perLabel(x))}</b><small>Fact. ${esc(x.id)}</small></td><td class="num" data-l="Costo total">${fmt0(t)}</td><td class="num" data-l="Var. vs. anterior">${vt}</td>
        <td class="pa">${cur ? `<span class="pcur">Actual</span>` : `<button type="button" class="peye" data-per="${esc(x.id)}" aria-label="Ver el período ${esc(perLabel(x))}" title="Ver este período">${ICON.eye}</button>`}</td></tr>`;
    }).join("")}</tbody></table></div>
    ${pers.length > 3 ? `<button type="button" class="dlink" id="d-hist">${S.dashAll ? "Ver solo los últimos 3" : `Ver historial completo (${pers.length})`}</button>` : ""}
    <p class="dfoot">La variación compara el total facturado con el período anterior.</p>`;

  // Alertas y novedades, armadas con el estado real del período
  const al = [];
  if (d.sin) al.push(["warn", `${d.sin} ${d.sin === 1 ? "patente" : "patentes"} sin asignación`, "Asignar el centro de costo que corresponde", "none"]);
  if (d.inc) al.push(["bad", `${d.inc} ${d.inc === 1 ? "patente" : "patentes"} con CC no habilitado`, "Reasignar a un centro de costo habilitado", "bad"]);
  const dif = Math.round((d.T - d.F) * 100) / 100;
  al.push(Math.abs(dif) < 0.005 ? ["ok", "El prorrateo cuadra con la factura", `Total distribuido ${fmt0(d.F)}, al centavo`, ""]
    : ["bad", "Diferencia entre factura y prorrateo", `$ ${fmt(Math.abs(dif))} (${fmtPct(d.F ? Math.abs(dif) / d.F : 0)})`, ""]);
  if (d.nuevas) al.push(["info", `${d.nuevas} ${d.nuevas === 1 ? "patente nueva" : "patentes nuevas"} en este período`, "No estaban en el período anterior", "new"]);
  if (d.sinZona && Object.keys(S.zonas).length) al.push(["info", `${d.sinZona} ${d.sinZona === 1 ? "patente" : "patentes"} sin zona`, "Actualizar la base de zonas", "zona"]);
  const bc = S.bases.centros_costo;
  if (bc) al.push(["info", `Base de CC: ${fmtInt(ccOn.length)} habilitados de ${fmtInt(S.ccBase.length)}`, bc.cargado_en ? `Actualizada el ${fechaAr(new Date(bc.cargado_en))}` : esc(bc.archivo || ""), ""]);
  if (p.creado_en) al.push(["ok", `Período ${perLabel(p)} cargado correctamente`, `Factura ${p.id} · ${fechaAr(new Date(p.creado_en))}`, ""]);
  $("d-alerts").innerHTML = `<ul class="alerts">${al.map(([t, h, s, go]) => {
    const inner = `<span class="ai">${ICON["a_" + t]}</span><span class="at"><b>${esc(h)}</b><small>${esc(s)}</small></span>${go ? ICON.chev : ""}`;
    return `<li class="al ${t}">${go ? `<button type="button" data-go="${go}">${inner}</button>` : `<div>${inner}</div>`}</li>`;
  }).join("")}</ul>`;
}
function goPat(go) {
  S.q = ""; $("q").value = ""; S.sel.clear(); S.zf = ""; S.fil = "all";
  if (go === "zona") S.zf = "__none"; else S.fil = go;
  setView("pat"); renderPat(); scrollTo(0, 0);
}
$("dashview").addEventListener("click", e => {
  const go = e.target.closest("[data-go]");
  if (go) { if (go.dataset.go === "cc") { setView("cc"); scrollTo(0, 0); } else goPat(go.dataset.go); return; }
  const b = e.target.closest(".brow");
  if (b) { const k = b.dataset.cc; S.open.add(k); setView("cc"); renderCC(); scrollTo(0, 0); return; }
  const pe = e.target.closest(".peye");
  if (pe) { S.actual = pe.dataset.per; localStorage.setItem("seg-periodo", S.actual); S.sel.clear(); S.open.clear(); reload({ soloPeriodo: true }); scrollTo(0, 0); return; }
  if (e.target.closest("#d-hist")) { S.dashAll = !S.dashAll; renderDash(); }
});

/* ================= Menú de CC (estilo autofiltro de Excel) =================
   Buscador arriba, lista con casilleros y Aceptar / Cancelar abajo. Solo
   ofrece CC habilitados. Un vehículo lleva un solo CC: tildar uno destilda
   el anterior. Teclado: ↑ ↓ recorren la lista, Espacio tilda, Enter acepta,
   Esc cancela. Doble clic en una opción la tilda y acepta. */
const PK = { open: false, btn: null, opt: null, items: [], chk: null, act: -1 };
const pop = $("ccpop"), popQ = $("ccpop-q"), popL = $("ccpop-l");
const mobile = () => matchMedia("(max-width:760px)").matches;

function hl(text, terms) {
  // Resalta los términos sin distinguir mayúsculas ni tildes
  if (!terms.length) return esc(text);
  let norm = "", map = [];
  [...text].forEach((ch, i) => { const n = normH(ch) || ch.toLowerCase(); for (const _ of n) map.push(i); norm += n; });
  const chars = [...text], mark = new Array(chars.length).fill(false);
  for (const t of terms) for (let k = norm.indexOf(t); k >= 0; k = norm.indexOf(t, k + 1)) for (let j = k; j < k + t.length; j++) mark[map[j]] = true;
  let out = "", on = false;
  chars.forEach((ch, i) => { if (mark[i] !== on) { out += mark[i] ? "<mark>" : "</mark>"; on = mark[i]; } out += esc(ch); });
  return out + (on ? "</mark>" : "");
}
function pickList() {
  const terms = normH(popQ.value).split(/\s+/).filter(Boolean);
  const found = terms.length ? ccOn.filter(c => terms.every(t => c.key.includes(t))) : ccOn;
  PK.items = (PK.opt.allowNone && !terms.length ? [{ code: "", d: "Quitar el centro de costo", none: true }] : []).concat(found);
  if (PK.act >= PK.items.length) PK.act = PK.items.length - 1;
  popL.innerHTML = !PK.items.length
    ? `<div class="ccpop-empty">Ningún CC habilitado coincide con «${esc(popQ.value.trim())}».</div>`
    : PK.items.map((c, i) => `<div class="ccopt${c.none ? " none" : ""}${c.code === PK.chk ? " on" : ""}${i === PK.act ? " act" : ""}" role="option" id="ccopt-${i}" data-i="${i}" aria-selected="${c.code === PK.chk}">
        <span class="box" aria-hidden="true"></span>${c.none ? `<span class="code">(Sin CC)</span><span class="d">${esc(c.d)}</span>` : `<span class="code">${hl(c.code, terms)}</span><span class="d">${hl(c.d, terms)}</span>`}</div>`).join("");
  popQ.setAttribute("aria-activedescendant", PK.act >= 0 ? "ccopt-" + PK.act : "");
  const n = found.length;
  $("ccpop-n").textContent = terms.length ? `${fmtInt(n)} de ${fmtInt(ccOn.length)} CC habilitados` : `${fmtInt(ccOn.length)} CC habilitados`;
  $("ccpop-ok").disabled = PK.chk === null;
}
function pickActive(i, scroll = true) {
  PK.act = i;
  popL.querySelectorAll(".ccopt").forEach(el => el.classList.toggle("act", +el.dataset.i === i));
  popQ.setAttribute("aria-activedescendant", i >= 0 ? "ccopt-" + i : "");
  if (scroll && i >= 0) popL.querySelector(`[data-i="${i}"]`)?.scrollIntoView({ block: "nearest" });
}
function pickCheck(i) {
  const c = PK.items[i]; if (!c) return;
  PK.chk = PK.chk === c.code ? null : c.code;   // tildar de nuevo lo destilda, como en Excel
  popL.querySelectorAll(".ccopt").forEach(el => { const on = PK.items[+el.dataset.i].code === PK.chk; el.classList.toggle("on", on); el.setAttribute("aria-selected", on); });
  $("ccpop-ok").disabled = PK.chk === null;
  pickActive(i, false);
}
function placePicker() {
  if (!PK.open) return;
  pop.classList.toggle("sheet", mobile()); $("ccpop-bg").hidden = !mobile();
  if (mobile()) { pop.style.cssText = ""; return; }
  const r = PK.btn.getBoundingClientRect();
  if (!PK.btn.isConnected || r.bottom < 0 || r.top > innerHeight) return closePicker(false);
  const w = Math.min(Math.max(r.width, 360), innerWidth - 16);
  const below = innerHeight - r.bottom - 12, above = r.top - 12, up = below < 280 && above > below;
  const maxH = Math.min(440, up ? above : below);
  pop.style.width = w + "px";
  pop.style.maxHeight = Math.max(maxH, 200) + "px";
  pop.style.left = Math.max(8, Math.min(r.left, innerWidth - w - 8)) + "px";
  pop.style.top = up ? "" : r.bottom + 4 + "px";
  pop.style.bottom = up ? innerHeight - r.top + 4 + "px" : "";
}
function openPicker(btn, opt) {
  if (PK.open && PK.btn === btn) return closePicker();   // segundo clic en el mismo botón: cierra
  if (PK.open) closePicker(false);
  if (!ccOn.length) return toast("No hay centros de costo habilitados en la base. Subí la base de CC con la columna «Habilitado».");
  Object.assign(PK, { open: true, btn, opt, chk: !opt.cur && opt.allowNone ? "" : ccHabilitado(opt.cur) ? opt.cur : null, act: -1 });
  $("ccpop-t").textContent = opt.title;
  const w = $("ccpop-warn");
  w.hidden = !opt.cur || ccHabilitado(opt.cur);
  if (!w.hidden) w.textContent = ccMap.has(opt.cur) ? `El CC actual (${opt.cur}) está inhabilitado. Elegí uno habilitado.` : `El CC actual (${opt.cur}) no existe en la base. Elegí uno habilitado.`;
  popQ.value = "";
  pop.hidden = false; btn.setAttribute("aria-expanded", "true");
  pickList();
  const i = PK.chk === null ? -1 : PK.items.findIndex(c => c.code === PK.chk);
  pickActive(i >= 0 ? i : (PK.items.length ? 0 : -1));
  placePicker();
  popQ.focus({ preventScroll: true });
}
function closePicker(refocus = true) {
  if (!PK.open) return;
  PK.open = false; pop.hidden = true; $("ccpop-bg").hidden = true;
  PK.btn?.setAttribute("aria-expanded", "false");
  if (refocus) (PK.opt.refocus?.() || (PK.btn.isConnected ? PK.btn : null))?.focus({ preventScroll: true });
}
async function acceptPicker() {
  if (!PK.open || PK.chk === null) return;
  const code = PK.chk, opt = PK.opt;
  closePicker(false);
  const done = opt.onPick(code);
  (opt.refocus?.())?.focus({ preventScroll: true });
  await done;
}

popQ.addEventListener("input", () => { PK.act = -1; pickList(); pickActive(PK.items.length ? 0 : -1); });
popQ.addEventListener("keydown", e => {
  const n = PK.items.length;
  if (e.key === "ArrowDown") { e.preventDefault(); if (n) pickActive(Math.min(n - 1, PK.act + 1)); }
  else if (e.key === "ArrowUp") { e.preventDefault(); if (n) pickActive(Math.max(0, PK.act - 1)); }
  else if (e.key === "Home" && e.ctrlKey) { e.preventDefault(); if (n) pickActive(0); }
  else if (e.key === "End" && e.ctrlKey) { e.preventDefault(); if (n) pickActive(n - 1); }
  else if (e.key === "PageDown") { e.preventDefault(); if (n) pickActive(Math.min(n - 1, PK.act + 8)); }
  else if (e.key === "PageUp") { e.preventDefault(); if (n) pickActive(Math.max(0, PK.act - 8)); }
  else if (e.key === " " && !popQ.value.endsWith(" ") && PK.act >= 0 && (e.ctrlKey || !popQ.value)) { e.preventDefault(); pickCheck(PK.act); }
  else if (e.key === "Enter") {
    e.preventDefault();
    // Enter tilda la opción marcada (o la única que quedó) y acepta
    const solo = PK.items.filter(c => !c.none);
    if (PK.act >= 0) { PK.chk = PK.items[PK.act].code; }
    else if (solo.length === 1) { PK.chk = solo[0].code; }
    acceptPicker();
  }
});
pop.addEventListener("keydown", e => {
  if (e.key === "Escape") { e.preventDefault(); closePicker(); }
  else if (e.key === "Tab") {
    // El foco queda dentro del menú mientras está abierto
    const f = [popQ, $("ccpop-ok"), $("ccpop-no"), $("ccpop-x")].filter(x => !x.disabled);
    const i = f.indexOf(document.activeElement);
    e.preventDefault(); f[(i + (e.shiftKey ? -1 : 1) + f.length) % f.length].focus();
  }
});
popL.addEventListener("mousedown", e => e.preventDefault());   // el foco sigue en el buscador
popL.addEventListener("click", e => { const o = e.target.closest(".ccopt"); if (o) pickCheck(+o.dataset.i); });
popL.addEventListener("dblclick", e => { const o = e.target.closest(".ccopt"); if (o) { PK.chk = PK.items[+o.dataset.i].code; acceptPicker(); } });
$("ccpop-ok").addEventListener("click", acceptPicker);
$("ccpop-no").addEventListener("click", () => closePicker());
$("ccpop-x").addEventListener("click", () => closePicker());
$("ccpop-bg").addEventListener("click", () => closePicker());
document.addEventListener("mousedown", e => { if (PK.open && !pop.contains(e.target) && !PK.btn.contains(e.target)) closePicker(false); });
addEventListener("scroll", placePicker, { capture: true, passive: true });
addEventListener("resize", placePicker);

/* ================= PDF de la factura =================
   El PDF se guarda en Supabase al cargar el período (seg_guardar_factura_pdf) y se
   muestra en un visor propio con pdf.js, que se ve igual en escritorio y celular.
   Si el período se cargó antes de guardar los PDF, el visor permite subirlo; se
   controla que sea la misma factura antes de guardarlo. */
const pdfCache = new Map();
const PV = { url: "", bytes: null, zoom: 1, token: 0, id: "" };
const toB64 = u8 => { let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };
const fromB64 = b => { const bin = atob(String(b).replace(/\s/g, "")); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i); return u8; };
const fmtKB = n => n >= 1048576 ? (n / 1048576).toLocaleString("es-AR", { maximumFractionDigits: 1 }) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB";

async function loadPdfInfo() {
  // null = la migración del PDF todavía no está en la base
  const { data, error } = await sb.rpc("seg_facturas_pdf_info");
  S.pdfs = error ? null : Object.fromEntries((data || []).map(x => [x.periodo_id, x]));
}
async function guardarFacturaPdf(id, file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes.length > 10 * 1048576) throw new Error("el PDF supera los 10 MB");
  await rpc("seg_guardar_factura_pdf", { p_periodo: id, p_archivo: file.name, p_pdf: toB64(bytes) });
  if (S.pdfs) S.pdfs[id] = { periodo_id: id, archivo: file.name, bytes: bytes.length, cargado_en: new Date().toISOString() };
  pdfCache.set(id, bytes);
}
function renderPdfBtn(p) {
  const has = !!(S.pdfs && S.pdfs[p.id]);
  const t = has ? `Ver el PDF de la factura ${p.id}` : `La factura ${p.id} no tiene el PDF guardado: tocá para subirlo`;
  ["i-pdf", "d-pdf"].forEach(id => {
    const b = $(id);
    b.classList.toggle("missing", !has);
    b.querySelector(".pdfdot").hidden = has;
    b.title = t; b.setAttribute("aria-label", t);
  });
}
function pdfMsg(html) { $("pdfv-body").innerHTML = `<div class="pdfv-msg">${html}</div>`; }
function pdfClear() {
  if (PV.url) URL.revokeObjectURL(PV.url);
  Object.assign(PV, { url: "", bytes: null });
  $("pdfv-dl").hidden = true; $("pdfv-open").hidden = true; $("pdfv-zm").hidden = true; $("pdfv-zp").hidden = true;
}
function pdfSinArchivo(p) {
  pdfMsg(`<img src="/factura-pdf.png" alt="" width="72" height="72">
    <b>Esta factura no tiene el PDF guardado</b>
    <p>Se cargó antes de que la aplicación guardara los PDF. Subí el PDF de la factura ${esc(p.id)} y queda guardado para todos.</p>
    <button type="button" class="btn primary" id="pdfv-up">Subir el PDF de la factura</button>`);
}
async function openPdf(e) {
  const p = P(); if (!p) return;
  PV.from = e?.currentTarget || $("i-pdf");
  const tok = ++PV.token; PV.id = p.id; PV.zoom = 1; pdfClear();
  const info = S.pdfs && S.pdfs[p.id];
  $("pdfv-t").textContent = "Factura " + p.id;
  $("pdfv-s").textContent = [isoToAr(p.factura_fecha), info?.archivo, info?.bytes && fmtKB(info.bytes)].filter(Boolean).join(" · ");
  $("pdfv").hidden = false; document.body.classList.add("pdfv-on");
  $("pdfv-x").focus();
  if (S.pdfs === null) return pdfMsg(`<b>Falta preparar la base de datos</b><p>Para guardar y ver el PDF de las facturas, ejecutá en Supabase la migración <code>20261114120000_factura_pdf.sql</code>.</p>`);
  if (!info) return pdfSinArchivo(p);
  pdfMsg(`<span class="pdfv-spin" aria-hidden="true"></span><b>Abriendo la factura…</b>`);
  try {
    let bytes = pdfCache.get(p.id);
    if (!bytes) {
      const r = await rpc("seg_factura_pdf", { p_periodo: p.id });
      if (tok !== PV.token) return;
      if (!r || !r.contenido) { delete S.pdfs[p.id]; renderPdfBtn(p); return pdfSinArchivo(p); }
      bytes = fromB64(r.contenido); pdfCache.set(p.id, bytes);
    }
    if (tok !== PV.token) return;
    await showPdf(bytes, info.archivo || `Factura_${p.id}.pdf`, tok);
  } catch (e) { if (tok === PV.token) pdfMsg(`<b>No se pudo abrir el PDF</b><p>${esc(errMsg(e))}</p>`); }
}
async function showPdf(bytes, archivo, tok) {
  PV.bytes = bytes;
  PV.url = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
  const dl = $("pdfv-dl"); dl.href = PV.url; dl.download = /\.pdf$/i.test(archivo) ? archivo : archivo + ".pdf"; dl.hidden = false;
  $("pdfv-open").href = PV.url; $("pdfv-open").hidden = false;
  $("pdfv-zm").hidden = false; $("pdfv-zp").hidden = false;
  await renderPages(tok);
}
async function renderPages(tok) {
  const body = $("pdfv-body");
  try {
    const { getDocument } = await loadPdfjs();
    const doc = await getDocument({ data: PV.bytes.slice() }).promise;
    if (tok !== PV.token) return;
    const avail = Math.max(240, body.clientWidth - 32), dpr = Math.min(window.devicePixelRatio || 1, 2.5);
    const frag = document.createDocumentFragment();
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const base = page.getViewport({ scale: 1 });
      const scale = Math.min(avail / base.width, 2) * PV.zoom;
      const vp = page.getViewport({ scale: scale * dpr });
      const c = document.createElement("canvas");
      c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
      c.style.width = Math.floor(vp.width / dpr) + "px"; c.style.height = Math.floor(vp.height / dpr) + "px";
      c.setAttribute("role", "img"); c.setAttribute("aria-label", `Página ${i} de ${doc.numPages} de la factura`);
      await page.render({ canvasContext: c.getContext("2d"), viewport: vp, canvas: c }).promise;
      if (tok !== PV.token) return;
      frag.appendChild(c);
    }
    body.innerHTML = ""; body.appendChild(frag);
    $("pdfv-zm").disabled = PV.zoom <= 0.5; $("pdfv-zp").disabled = PV.zoom >= 3;
  } catch (e) { if (tok === PV.token) pdfMsg(`<b>No se pudo mostrar el PDF</b><p>Probá con «Descargar». ${esc(errMsg(e))}</p>`); }
}
function closePdf() {
  if ($("pdfv").hidden) return;
  PV.token++; pdfClear();
  $("pdfv").hidden = true; document.body.classList.remove("pdfv-on"); $("pdfv-body").innerHTML = "";
  (PV.from && PV.from.offsetParent ? PV.from : $("i-pdf")).focus({ preventScroll: true });
}
$("i-pdf").addEventListener("click", openPdf);
$("d-pdf").addEventListener("click", openPdf);
$("pdfv-x").addEventListener("click", closePdf);
$("pdfv").addEventListener("click", e => { if (e.target === $("pdfv")) closePdf(); });
document.addEventListener("keydown", e => {
  if ($("pdfv").hidden) return;
  if (e.key === "Escape") { e.preventDefault(); closePdf(); }
  else if (e.key === "Tab") {
    // El foco queda dentro del visor
    const f = [...$("pdfv").querySelectorAll("button,a[href],[tabindex='0']")].filter(x => !x.hidden && !x.disabled && x.offsetParent);
    const i = f.indexOf(document.activeElement);
    e.preventDefault(); f[(i + (e.shiftKey ? -1 : 1) + f.length) % f.length]?.focus();
  }
});
const zoom = k => { PV.zoom = Math.min(3, Math.max(0.5, PV.zoom * k)); renderPages(PV.token); };
$("pdfv-zp").addEventListener("click", () => zoom(1.25));
$("pdfv-zm").addEventListener("click", () => zoom(0.8));
$("pdfv-body").addEventListener("click", e => { if (e.target.closest("#pdfv-up")) $("pdfv-file").click(); });
$("pdfv-file").addEventListener("change", async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  const p = P(); if (!p || p.id !== PV.id) return;
  const tok = PV.token;
  pdfMsg(`<span class="pdfv-spin" aria-hidden="true"></span><b>Revisando el PDF…</b>`);
  try {
    // Antes de guardarlo se controla que sea la factura de este período
    const { getDocument } = await loadPdfjs();
    const bytes = new Uint8Array(await f.arrayBuffer());
    const doc = await getDocument({ data: bytes.slice() }).promise;
    const fa = parseFacturaLines(await pageText(doc, 1));
    if (fa.numero !== p.id) {
      if (tok === PV.token) pdfSinArchivo(p);
      return toast(fa.numero ? `Ese PDF es la factura ${fa.numero}, no la ${p.id}.` : `«${f.name}» no parece ser la factura ${p.id}.`);
    }
    await guardarFacturaPdf(p.id, f);
    renderPdfBtn(p);
    if (tok !== PV.token) return;
    $("pdfv-s").textContent = [isoToAr(p.factura_fecha), f.name, fmtKB(bytes.length)].filter(Boolean).join(" · ");
    toast("PDF de la factura guardado.");
    await showPdf(bytes, f.name, tok);
    $("pdfv-body").focus({ preventScroll: true });
  } catch (err) {
    if (tok === PV.token) pdfSinArchivo(p);
    toast("No se guardó el PDF: " + errMsg(err));
  }
});
let pdfResizeT;
addEventListener("resize", () => { if (!$("pdfv").hidden && PV.bytes) { clearTimeout(pdfResizeT); pdfResizeT = setTimeout(() => renderPages(PV.token), 200); } });

/* ================= Excel: lectura ================= */
async function readSheet(file) { const wb = XLSX.read(await file.arrayBuffer(), { type: "array" }); return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: false }); }
const isCCHeader = c => ["cc", "centro de costo", "centro de costos"].includes(normH(c));

$("imp").addEventListener("click", () => $("impf").click());
$("impf").addEventListener("change", async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  try {
    const aoa = await readSheet(f);
    const hi = aoa.findIndex(r => r.some(c => normH(c) === "patente") && r.some(isCCHeader));
    if (hi < 0) return toast("No encontré columnas «Patente» y «CC» en la primera hoja.");
    const h = aoa[hi], ip = h.findIndex(c => normH(c) === "patente"), ic = h.findIndex(isCCHeader);
    const known = new Set(rows.map(r => r.patente)); const by = {};
    let skip = 0; const off = new Set();
    for (const r of aoa.slice(hi + 1)) {
      const pt = normPat(r[ip]), c = codeFrom(r[ic]); if (!pt) continue;
      if (!known.has(pt) || !c) { skip++; continue; }
      if (!ccHabilitado(c)) { off.add(c); skip++; continue; }
      (by[c] = by[c] || []).push(pt);
    }
    let ok = 0;
    for (const [c, ps] of Object.entries(by)) if (await assign(ps, c)) ok += ps.length;
    toast(`Importadas ${ok} asignaciones` + (skip ? `; ${skip} filas omitidas (patente fuera de la póliza, sin CC o CC no habilitado${off.size ? ": " + [...off].slice(0, 6).join(", ") + (off.size > 6 ? "…" : "") : ""}).` : "."));
  } catch (err) { toast("No se pudo leer el archivo. Usá un .xlsx con columnas Patente y CC."); }
});

async function parseZonas(f) {
  try {
    const aoa = await readSheet(f);
    const hi = aoa.findIndex(r => r.some(c => normH(c) === "patente") && r.some(c => normH(c) === "zona"));
    if (hi < 0) return { error: "No encontré columnas «Patente» y «Zona» en la primera hoja." };
    const h = aoa[hi].map(normH), ip = h.indexOf("patente"), iz = h.indexOf("zona");
    const map = {}, conflictos = [];
    for (const r of aoa.slice(hi + 1)) {
      const pt = normPat(r[ip]), z = String(r[iz]).trim().replace(/\s+/g, " ");
      if (!pt || !z) continue;
      if (map[pt] && map[pt] !== z && !conflictos.includes(pt)) conflictos.push(pt);
      map[pt] = z;
    }
    if (!Object.keys(map).length) return { error: "La hoja de zonas no tiene patentes con zona." };
    return { map, conflictos, archivo: f.name };
  } catch (e) { return { error: "No se pudo leer el Excel de zonas." }; }
}
async function guardarZonas(z) {
  await rpc("seg_reemplazar_zonas", { p_items: Object.entries(z.map).map(([patente, zona]) => ({ patente, zona })), p_archivo: z.archivo, p_conflictos: z.conflictos });
}
$("impz").addEventListener("click", () => $("impzf").click());
$("impzf").addEventListener("change", async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  const z = await parseZonas(f);
  if (z.error) return toast(z.error);
  try {
    await guardarZonas(z); await reload();
    const con = rows.filter(r => r.zona).length;
    toast(`Zonas actualizadas: ${con} de ${rows.length} patentes tienen zona.` + (z.conflictos.length ? ` Revisá ${z.conflictos.join(", ")}: figura con más de una zona.` : ""));
  } catch (err) { toast("No se guardaron las zonas: " + errMsg(err)); }
});

/* ================= Carga de período ================= */
function showUpload(on, forced) {
  $("upload").hidden = !on; $("upload").dataset.forced = forced ? "1" : "";
  $("cancel").hidden = !P();
  if (on) renderUpload();
}
$("newper").addEventListener("click", () => { S.up = { poliza: null, factura: null, cc: null, zona: null }; showUpload(true, true); $("upload").scrollIntoView({ behavior: "smooth" }); });
$("cancel").addEventListener("click", () => { showUpload(false); renderAll(); });
const drop = $("drop");
drop.addEventListener("click", () => $("files").click());
drop.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("files").click(); } });
["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", e => handleFiles([...e.dataTransfer.files]));
$("files").addEventListener("change", e => { handleFiles([...e.target.files]); e.target.value = ""; });

function slot(name, st, cls, fn, info) {
  const el = $("s-" + name); const s = el.querySelector(".st"); s.textContent = st; s.className = "st " + cls;
  if (fn !== undefined) el.querySelector(".fn").textContent = fn;
  if (info !== undefined) el.querySelector(".info").innerHTML = info;
}
function renderUpload() {
  const u = S.up, bc = S.bases.centros_costo, bz = S.bases.zonas;
  if (!u.poliza) slot("poliza", "Falta", "wait", "", "Las páginas de cada certificado con patente, marca y valor.");
  if (!u.factura) slot("factura", "Falta", "wait", "", "Factura de la renovación de la póliza.");
  if (!u.cc) {
    if (S.ccBase.length) slot("cc", "Guardada", "ok", bc?.archivo || "", `Se usa la base guardada (${fmtInt(S.ccBase.length)} CC). Subí una nueva si cambió.`);
    else slot("cc", "Falta", "wait", "", "Excel con columnas CC, Descripción y Habilitado.");
  }
  if (!u.zona) {
    if (Object.keys(S.zonas).length) slot("zona", "Guardada", "ok", bz?.archivo || "", `Se usan las zonas guardadas (${fmtInt(Object.keys(S.zonas).length)} patentes). Subí una nueva si cambió.`);
    else slot("zona", "Opcional", "wait", "", "Excel con columnas Patente y Zona.");
  }
  renderChecks();
}
async function handleFiles(files) {
  const pdfs = [];
  let getDocument = null;
  if (files.some(f => /\.pdf$/i.test(f.name))) {
    try { ({ getDocument } = await loadPdfjs()); }
    catch (e) { return toast("No se pudo cargar el lector de PDF. Recargá la página e intentá de nuevo."); }
  }
  for (const f of files) {
    if (/\.(xlsx|xls)$/i.test(f.name)) {
      let kind = "cc";
      try { const aoa = await readSheet(f); if (aoa.slice(0, 10).some(r => r.some(c => normH(c) === "patente") && r.some(c => normH(c) === "zona"))) kind = "zona"; } catch (e) { }
      if (kind === "zona") await loadZona(f); else await loadCC(f);
    } else if (/\.pdf$/i.test(f.name)) {
      try {
        const pdf = await getDocument({ data: new Uint8Array(await f.arrayBuffer()) }).promise;
        const first = (await pageText(pdf, 1)).map(l => l.text).join("\n");
        if (/Póliza para el seguro/i.test(first) || /TOTAL CERTIFICADOS/.test(first)) pdfs.push(["poliza", f, pdf]);
        else if (/FACTURA/.test(first) && /C\.A\.E/.test(first)) pdfs.unshift(["factura", f, pdf]);
        else toast(`«${f.name}» no parece ser ni la póliza ni la factura de La Segunda.`);
      } catch (err) { toast(`No se pudo abrir «${f.name}».`); }
    } else toast(`«${f.name}» no es PDF ni Excel.`);
  }
  for (const [kind, f, pdf] of pdfs) {
    if (kind === "factura") {
      const fa = parseFacturaLines(await pageText(pdf, 1));
      if (fa.neto == null || fa.total == null) { slot("factura", "Error", "bad", f.name, "No encontré los importes de la factura."); continue; }
      S.up.factura = { archivo: f.name, file: f, ...fa };
      slot("factura", "Lista", "ok", f.name, `N° ${esc(fa.numero)} del ${esc(fa.fecha)}<br>Total $ ${fmt(fa.total)}`);
    } else {
      slot("poliza", "Leyendo", "busy", f.name, "Abriendo…"); const mini = $("s-poliza").querySelector(".mini"); mini.hidden = false;
      try {
        const res = await parsePoliza(pdf, (p, n, v) => { mini.firstElementChild.style.width = (p / n * 100) + "%"; slot("poliza", "Leyendo", "busy", undefined, `Página ${fmtInt(p)} de ${fmtInt(n)} · ${v} vehículos`); });
        mini.hidden = true;
        if (!res.vehiculos.length) { slot("poliza", "Error", "bad", f.name, "No encontré certificados con datos de vehículos."); continue; }
        S.up.poliza = { archivo: f.name, ...res };
        slot("poliza", "Lista", "ok", f.name, `Póliza ${esc(res.meta.poliza || "")} · ${res.vehiculos.length} vehículos<br>Vigencia ${esc(res.meta.desde || "")} al ${esc(res.meta.hasta || "")}`);
      } catch (err) { mini.hidden = true; slot("poliza", "Error", "bad", f.name, "No se pudo leer la póliza."); }
    }
    renderChecks();
  }
  renderUpload();
}
async function loadZona(f) {
  const z = await parseZonas(f);
  if (z.error) { slot("zona", "Error", "bad", f.name, esc(z.error)); return; }
  S.up.zona = z;
  slot("zona", "Lista", "ok", f.name, `${fmtInt(Object.keys(z.map).length)} patentes · ${new Set(Object.values(z.map)).size} zonas`);
}
async function loadCC(f) {
  try {
    const aoa = await readSheet(f), hi = aoa.findIndex(r => r.some(c => normH(c) === "cc"));
    if (hi < 0) { slot("cc", "Error", "bad", f.name, "No encontré la columna «CC» en la primera hoja."); return; }
    const h = aoa[hi].map(normH), ic = h.indexOf("cc"), id = h.findIndex(c => c.startsWith("descrip"));
    // «Habilitado» va en la columna C; si el título no está, se toma esa columna.
    let ih = h.findIndex(c => c.startsWith("habilit"));
    if (ih < 0 && ic !== 2 && id !== 2 && aoa[hi].length > 2) ih = 2;
    if (ih < 0) { slot("cc", "Error", "bad", f.name, "No encontré la columna «Habilitado» (columna C)."); return; }
    const list = aoa.slice(hi + 1).filter(r => String(r[ic]).trim()).map(r => {
      const hv = normH(r[ih]);
      return { codigo: String(r[ic]).trim(), descripcion: id < 0 ? "" : String(r[id]).trim(), habilitado: ["verdadero", "true", "si", "1", "x", "yes"].includes(hv) };
    });
    if (!list.length) { slot("cc", "Error", "bad", f.name, "La hoja no tiene centros de costo."); return; }
    S.up.cc = { archivo: f.name, list };
    slot("cc", "Lista", "ok", f.name, `${fmtInt(list.length)} centros de costo · ${fmtInt(list.filter(x => x.habilitado).length)} habilitados`);
  } catch (e) { slot("cc", "Error", "bad", f.name, "No se pudo leer el Excel."); }
}
function renderChecks() {
  const u = S.up, out = [];
  const pol = u.poliza, fa = u.factura, hayBase = u.cc || S.ccBase.length;
  if (pol) {
    const v = pol.vehiculos, n = pol.meta.totalCert;
    out.push(n && n !== v.length ? ["bad", `Encontré ${v.length} vehículos pero la póliza declara ${n} certificados.`] : ["", `${v.length} vehículos extraídos${n ? `, coincide con los ${n} certificados declarados` : ""}.`]);
    const seen = {}, dup = []; v.forEach(x => { if (seen[x.patente]) dup.push(x.patente); seen[x.patente] = 1; });
    if (dup.length) out.push(["warn", `Patentes repetidas: ${dup.join(", ")}.`]);
    const sinPrem = v.filter(x => !x.premio).length; if (sinPrem) out.push(["warn", `${sinPrem} vehículos sin premio; no recibirán importe en el prorrateo.`]);
    const sinPat = v.filter(x => !x.patente).length; if (sinPat) out.push(["bad", `${sinPat} certificados sin patente.`]);
  }
  if (fa) {
    const s = COMP.reduce((a, c) => a + (fa[c] || 0), 0);
    out.push(Math.abs(s - fa.total) < 0.01 ? ["", `Factura consistente: neto + IVA + percepción + otros = $ ${fmt(fa.total)}.`] : ["bad", `Los conceptos de la factura suman $ ${fmt(s)} y el total dice $ ${fmt(fa.total)}.`]);
  }
  if (pol && fa) {
    if (pol.meta.premioFinal != null) out.push(Math.abs(pol.meta.premioFinal - fa.total) < 0.01 ? ["", "El total de la factura coincide con el premio final de la póliza."] : ["warn", `La factura ($ ${fmt(fa.total)}) no coincide con el premio final de la póliza ($ ${fmt(pol.meta.premioFinal)}).`]);
    const pn = (pol.meta.poliza || "").replace(/\D/g, "");
    if (pn && !(fa.detalle || "").replace(/\D/g, "").includes(pn)) out.push(["warn", `El detalle de la factura no menciona la póliza ${pol.meta.poliza}.`]);
    const zmap = u.zona ? u.zona.map : S.zonas;
    if (Object.keys(zmap).length) {
      const sin = pol.vehiculos.filter(x => !zmap[x.patente]).map(x => x.patente);
      out.push(sin.length ? ["warn", `${pol.vehiculos.length - sin.length} patentes con zona; ${sin.length} sin zona en la base (${sin.slice(0, 8).join(", ")}${sin.length > 8 ? "…" : ""}).`] : ["", "Todas las patentes tienen zona."]);
    }
    const conf = u.zona ? u.zona.conflictos : (S.bases.zonas?.detalle?.conflictos || []);
    if (conf.length) out.push(["warn", `Con más de una zona en la base (se toma la última): ${conf.join(", ")}.`]);
    if (S.asignadas.size) { const prev = pol.vehiculos.filter(x => S.asignadas.has(x.patente)).length; out.push(["", `${prev} patentes toman el CC del período anterior; ${pol.vehiculos.length - prev} quedan para asignar.`]); }
    if (S.periodos.some(p => p.id === fa.numero)) out.push(["warn", "Esta factura ya estaba cargada; al guardar se reemplazan sus datos y se conservan sus asignaciones."]);
  }
  $("checks").innerHTML = out.map(([c, t]) => `<li class="${c}">${esc(t)}</li>`).join("");
  const ready = pol && fa && hayBase && !out.some(x => x[0] === "bad");
  $("save").disabled = !ready;
  $("savenote").textContent = ready ? "" : !pol || !fa || !hayBase ? `Falta: ${[!pol && "póliza", !fa && "factura", !hayBase && "base de CC"].filter(Boolean).join(", ")}.` : "Corregí los errores marcados en rojo.";
}
$("save").addEventListener("click", async () => {
  const { poliza: pol, factura: fa, cc, zona } = S.up, btn = $("save");
  busy(btn, true, "Guardando…");
  try {
    if (cc) await rpc("seg_reemplazar_centros_costo", { p_items: cc.list, p_archivo: cc.archivo });
    if (zona) await guardarZonas(zona);
    const res = await rpc("seg_guardar_periodo", {
      p_periodo: {
        id: fa.numero, poliza: pol.meta.poliza, tomador: pol.meta.tomador, vigencia_desde: pol.meta.desde, vigencia_hasta: pol.meta.hasta,
        total_certificados: pol.meta.totalCert, premio_final: pol.meta.premioFinal, factura_fecha: fa.fecha, factura_detalle: fa.detalle, cae: fa.cae,
        neto: fa.neto, iva: fa.iva, percepcion: fa.perc, otros_impuestos: fa.otros, total: fa.total, archivo_poliza: pol.archivo, archivo_factura: fa.archivo
      },
      p_vehiculos: pol.vehiculos.map(v => ({ patente: v.patente, certificado: v.cert, marca: v.marca, anio: v.anio, valor: v.valor, premio: v.premio }))
    });
    S.actual = res.id; localStorage.setItem("seg-periodo", S.actual);
    // El PDF de la factura se guarda aparte; si falla, el período ya quedó guardado
    let pdfErr = "";
    try { await guardarFacturaPdf(res.id, fa.file); } catch (e) { pdfErr = errMsg(e); }
    S.up = { poliza: null, factura: null, cc: null, zona: null };
    showUpload(false); S.open.clear(); setView("dash");
    await reload();
    const pend = res.vehiculos - res.con_cc;
    toast((pend ? `Período guardado. ${pend} patentes necesitan CC.` : "Período guardado. Todas las patentes tienen CC.")
      + (pdfErr ? ` No se guardó el PDF de la factura (${pdfErr}); podés subirlo desde el ícono de la factura.` : ""));
  } catch (e) { toast("No se pudo guardar el período: " + errMsg(e)); }
  busy(btn, false, "Guardar y ver totales por CC");
});

/* ================= Excel: exportación ================= */
function exportXlsx() {
  const p = P();
  const baseTxt = "Premio del certificado";
  const money = '#,##0.00';
  const g = groupByCC();
  const sh = [[`Factura ${p.id} del ${isoToAr(p.factura_fecha)} — Póliza ${p.poliza} (${p.tomador || ""}) — Vigencia ${isoToAr(p.vigencia_desde)} al ${isoToAr(p.vigencia_hasta)}`], [`Base de prorrateo: ${baseTxt}`], [], ["CC", "Descripción", "Zonas", "Patentes", "Neto", "IVA", "Percepción IVA", "Otros impuestos", "Total"]];
  g.forEach(x => { const c = ccMap.get(x.cc); sh.push([x.cc || "(sin CC)", x.cc ? (c ? c.d : "No existe en base") : "Sin centro de costo", x.cc ? [...new Set(x.pats.map(r => r.zona).filter(Boolean))].join(", ") : "", x.n, x.neto, x.iva, x.perc, x.otros, x.total]); });
  const a = 5, b = sh.length;
  sh.push(["TOTAL", "", "", ...["D", "E", "F", "G", "H", "I"].map(C => ({ f: `SUM(${C}${a}:${C}${b})` }))]);
  const ws1 = XLSX.utils.aoa_to_sheet(sh);
  for (let R = a; R <= b + 1; R++) ["E", "F", "G", "H", "I"].forEach(C => { const c = ws1[C + R]; if (c) c.z = money; });
  ws1["!cols"] = [12, 34, 30, 9, 16, 14, 14, 14, 16].map(w => ({ wch: w }));

  const head = [["Cert. N°", "Patente", "Marca completa", "Zona", "Año", "Valor del vehículo", "Premio certificado", "% prorrateo", "Neto", "IVA", "Percepción IVA", "Otros impuestos", "Total", "CC", "Descripción CC", "Estado CC"]];
  const body = rows.slice().sort((x, y) => (x.cert || 0) - (y.cert || 0)).map(r => { const c = ccMap.get(r.cc); return [r.cert, r.patente, r.marca, r.zona || "", Number(r.anio) || r.anio, r.valor, r.premio, r.pct, r.neto, r.iva, r.perc, r.otros, r.total, r.cc, c ? c.d : "", !r.cc ? "Sin asignar" : !c ? "No existe en base" : c.on ? "Habilitado" : "Inhabilitado"]; });
  const last = body.length + 1;
  const tot = ["", "TOTAL", "", "", "", ...["F", "G", "H", "I", "J", "K", "L", "M"].map(C => ({ f: `SUM(${C}2:${C}${last})` }))];
  const ws2 = XLSX.utils.aoa_to_sheet([...head, ...body, tot]);
  for (let R = 2; R <= last + 1; R++) { ["F", "G", "I", "J", "K", "L", "M"].forEach(C => { const c = ws2[C + R]; if (c) c.z = money; }); const q = ws2["H" + R]; if (q) q.z = "0.000%"; }
  ws2["!cols"] = [8, 10, 34, 22, 6, 16, 16, 11, 16, 14, 14, 14, 16, 10, 32, 14].map(w => ({ wch: w }));
  ws2["!autofilter"] = { ref: `A1:P${last}` };

  const ws3 = XLSX.utils.aoa_to_sheet([["Concepto", "Importe"], ["Neto (subtotal)", +p.neto], ["IVA", +p.iva], ["Percepción IVA", +p.percepcion], ["Otros impuestos", +p.otros_impuestos], ["Total", { f: "SUM(B2:B5)" }], [], ["Factura", p.id], ["Fecha", isoToAr(p.factura_fecha)], ["CAE", p.cae], ["Póliza", p.poliza]]);
  for (let R = 2; R <= 6; R++) { const c = ws3["B" + R]; if (c) c.z = money; } ws3["!cols"] = [{ wch: 20 }, { wch: 22 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "Totales por CC");
  XLSX.utils.book_append_sheet(wb, ws2, "Detalle por patente");
  XLSX.utils.book_append_sheet(wb, ws3, "Factura");
  XLSX.writeFile(wb, `Polizas_flota_vehicular_${p.id.replace(/[^\w-]/g, "")}.xlsx`);
}
$("exp").addEventListener("click", () => { const m = rows.filter(r => !r.cc).length; if (m) toast(`Exportando con ${m} patentes sin CC; figuran como «Sin CC».`); exportXlsx(); });

/* ================= Ingreso ================= */
const LS_USER = "seg-usuario";
function screen(name) {
  $("loading").hidden = name !== "loading"; $("login").hidden = name !== "login"; $("app").hidden = name !== "app";
}
function salir(msg) {
  localStorage.removeItem(LS_USER);
  conectar(null);
  $("lemail").value = ""; $("lerr").textContent = msg || "";
  screen("login"); $("lemail").focus();
}
async function validar(email) {
  conectar(null);
  const { data, error } = await sb.rpc("seg_ingresar", { p_email: email });
  if (error) throw error;
  return data; // {email, nombre} o null
}
async function entrar(u) {
  localStorage.setItem(LS_USER, JSON.stringify(u));
  conectar(u.email);
  $("uemail").textContent = u.nombre ? `${u.nombre} (${u.email})` : u.email;
  screen("loading");
  try {
    await loadMaestros(); await loadPeriodo();
    screen("app"); renderAll();
  } catch (e) { $("loading").textContent = "No se pudo conectar con la base de datos: " + errMsg(e); }
}
$("loginf").addEventListener("submit", async e => {
  e.preventDefault();
  const email = $("lemail").value.trim().toLowerCase(), btn = $("lbtn");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { $("lerr").textContent = "Escribí un correo válido."; return; }
  $("lerr").textContent = ""; busy(btn, true, "Ingresando…");
  try {
    const u = await validar(email);
    if (u) await entrar(u);
    else $("lerr").textContent = "Ese correo no está habilitado. Pedí que lo agreguen a la lista de usuarios.";
  } catch (err) { $("lerr").textContent = "No se pudo conectar con la base de datos: " + errMsg(err); }
  busy(btn, false, "Ingresar");
});
$("logout").addEventListener("click", () => salir());

/* ================= Inicio ================= */
async function start() {
  let guardado = null;
  try { guardado = JSON.parse(localStorage.getItem(LS_USER) || "null"); } catch (e) { }
  if (!guardado?.email) return salir();
  screen("loading");
  try {
    const u = await validar(guardado.email);
    if (u) await entrar(u); else salir("Tu correo ya no está habilitado.");
  } catch (e) { $("loading").textContent = "No se pudo conectar con la base de datos: " + errMsg(e); }
}
if (!SUPABASE_URL || !SUPABASE_KEY) {
  $("loading").textContent = "Falta configurar VITE_SUPABASE_URL y VITE_SUPABASE_PUBLISHABLE_KEY.";
} else {
  start();
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible" && !$("app").hidden && $("upload").hidden) reload(); });
}

/* Encabezado fijo: su altura ubica la barra de selección y los títulos de la tabla justo debajo.
   En tablet y celular el encabezado se desliza hacia arriba hasta dejar visibles solo las pestañas (--hpin). */
function setBulkH() {
  const b = $("bulk"), on = b.classList.contains("on"), fixed = getComputedStyle(b).position === "fixed";
  document.documentElement.style.setProperty("--bh", on && !fixed ? b.offsetHeight + "px" : "0px");
  document.documentElement.style.setProperty("--bfix", on && fixed ? b.offsetHeight + "px" : "0px");
}
(() => {
  const h = document.querySelector("header.top"), tabs = $("tabs");
  if (!h) return;
  const setH = () => {
    const r = document.documentElement.style;
    r.setProperty("--hh", h.offsetHeight + "px");
    r.setProperty("--hpin", (tabs.hidden ? h.offsetHeight : Math.max(0, tabs.offsetTop - 8)) + "px");
    setBulkH();
  };
  setH();
  if ("ResizeObserver" in window) { const ro = new ResizeObserver(setH); ro.observe(h); ro.observe($("bulk")); }
  addEventListener("resize", setH);
})();
