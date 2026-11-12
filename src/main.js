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
  view: "cc", q: "", fil: "all", zf: "", showOff: false, sel: new Set(), sort: { k: "cert", dir: 1 }, open: new Set(),
  up: { poliza: null, factura: null, cc: null, zona: null }
};
let ccMap = new Map();
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
  const dl = $("cclist"); dl.innerHTML = "";
  const frag = document.createDocumentFragment();
  for (const [code, c] of ccMap) {
    if (!c.on && !S.showOff) continue;
    const o = document.createElement("option"); o.value = code + " · " + c.d + (c.on ? "" : " (inhabilitado)"); frag.appendChild(o);
  }
  dl.appendChild(frag);
}
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
  if (!p) { $("main").hidden = true; showUpload(true); $("hsub").textContent = "Cargá la póliza, la factura y la base de CC para ver cuánto le toca a cada centro de costo."; return; }
  if ($("upload").dataset.forced !== "1") $("upload").hidden = true;
  $("main").hidden = false;
  $("hsub").textContent = `Póliza ${p.poliza || ""} · ${p.tomador || ""} · Vigencia ${isoToAr(p.vigencia_desde)} al ${isoToAr(p.vigencia_hasta)} · ${rows.length} vehículos`;
  renderInvoice(p); renderProgress(); renderCC(); renderPat();
}
function renderInvoice(p) {
  $("i-num").textContent = "Factura " + p.id;
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
      <td class="ccc"><div class="cc ${st.cls}"><input class="ccin" list="cclist" value="${esc(r.cc)}" placeholder="Elegí o escribí el CC" aria-label="Centro de costo de ${esc(r.patente)}"><small title="${esc(st.txt)}">${esc(st.txt)}</small></div></td>
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
  S.view = v; $("ccview").hidden = v !== "cc"; $("patview").hidden = v !== "pat";
  $("t-cc").setAttribute("aria-selected", v === "cc"); $("t-pat").setAttribute("aria-selected", v === "pat");
}
function refreshViews() { const y = scrollY; renderProgress(); renderCC(); renderPat(); scrollTo(0, y); }

/* ================= Asignación ================= */
async function assign(pats, code) {
  const p = P(); if (!p || !pats.length) return false;
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
  if (e.target.classList.contains("rs")) { e.target.checked ? S.sel.add(pt) : S.sel.delete(pt); tr.classList.toggle("sel", e.target.checked); syncBulk(); return; }
  if (e.target.classList.contains("ccin")) assign([pt], codeFrom(e.target.value));
});
function syncBulk() {
  const n = S.sel.size; $("bulk").classList.toggle("on", n > 0); document.body.classList.toggle("bulk-on", n > 0); setBulkH(); $("bulk-n").textContent = n === 1 ? "1 seleccionada" : `${n} seleccionadas`;
  const vis = [...$("tb").querySelectorAll("tr[data-p]")].map(t => t.dataset.p), all = $("all");
  all.checked = vis.length > 0 && vis.every(p => S.sel.has(p)); all.indeterminate = !all.checked && vis.some(p => S.sel.has(p));
}
$("all").addEventListener("change", e => { [...$("tb").querySelectorAll("tr[data-p]")].forEach(t => e.target.checked ? S.sel.add(t.dataset.p) : S.sel.delete(t.dataset.p)); renderPat(); });
$("bulk-ok").addEventListener("click", async () => {
  const code = codeFrom($("bulk-cc").value);
  if (!code) return toast("Elegí un centro de costo para asignar.");
  if (!ccMap.has(code)) return toast(`El CC ${code} no está en la base.`);
  const pats = [...S.sel], btn = $("bulk-ok"); busy(btn, true);
  if (await assign(pats, code)) { toast(`CC ${code} asignado a ${pats.length} patentes.`); S.sel.clear(); $("bulk-cc").value = ""; refreshViews(); }
  busy(btn, false);
});
$("bulk-clear").addEventListener("click", async () => { const pats = [...S.sel]; if (await assign(pats, "")) { S.sel.clear(); refreshViews(); } });
$("bulk-x").addEventListener("click", () => { S.sel.clear(); renderPat(); });
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
$("showoff").addEventListener("change", e => { S.showOff = e.target.checked; buildCC(); });
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
    let skip = 0;
    for (const r of aoa.slice(hi + 1)) { const pt = normPat(r[ip]), c = codeFrom(r[ic]); if (!pt) continue; if (known.has(pt) && c) (by[c] = by[c] || []).push(pt); else skip++; }
    let ok = 0;
    for (const [c, ps] of Object.entries(by)) if (await assign(ps, c)) ok += ps.length;
    toast(`Importadas ${ok} asignaciones` + (skip ? `; ${skip} filas sin patente de la póliza o sin CC.` : "."));
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
      S.up.factura = { archivo: f.name, ...fa };
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
    const h = aoa[hi].map(normH), ic = h.indexOf("cc"), id = h.findIndex(c => c.startsWith("descrip")), ih = h.findIndex(c => c.startsWith("habilit"));
    const list = aoa.slice(hi + 1).filter(r => String(r[ic]).trim()).map(r => {
      const hv = ih < 0 ? "verdadero" : normH(r[ih]);
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
    S.up = { poliza: null, factura: null, cc: null, zona: null };
    showUpload(false); S.open.clear(); setView("cc");
    await reload();
    const pend = res.vehiculos - res.con_cc;
    toast(pend ? `Período guardado. ${pend} patentes necesitan CC.` : "Período guardado. Todas las patentes tienen CC.");
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
