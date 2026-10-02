const C = window.CONFIG;
const $ = id => document.getElementById(id);
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const GRUPOS = ["KIA", "SDA", "ADP"];
const CANON = {
  KIA: ["Oeste","Tobalaba","Maipú","Arauco","Gran Avenida","Quilín","Melipilla"],
  SDA: ["SUBARU Plaza Oeste","SUBARU Plaza Tobalaba","SUBARU Plaza Sur","SUBARU Plaza Vespucio","DFSK Plaza Oeste","DFSK Tobalaba","DFSK Gran Avenida","DFSK Plaza Sur","DFSK Arauco Maipu","DFSK Melipilla","SUZUKI Bavaro","GWM Bavaro"],
  ADP: ["ADP Plaza Oeste","ADP Plaza Tobalaba","LANDKING Gran Avenida","ADP Maipú"]
};
const CIERRE = 12 * 60, BARRIDO = 18 * 60; // minutos del día (Santiago)
let DATA = [], chart = null, msal = null, tick = 0;

// ---------- Utilidades ----------
const normName = s => String(s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim();
const CANON_MAP = new Map(GRUPOS.flatMap(g => CANON[g].map(n => [g + "|" + normName(n), n])));
const pad = n => String(n).padStart(2, "0");
const ddmm = iso => iso.slice(8, 10) + "-" + iso.slice(5, 7) + "-" + iso.slice(0, 4);
const pct = (a, n) => n ? Math.round(a / n * 100) : null;
const fmtPct = p => p === null ? "–" : p + "%";

function nowSt() {
  const o = Object.fromEntries(new Intl.DateTimeFormat("en-CA", { timeZone: "America/Santiago", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23" })
    .formatToParts(new Date()).map(x => [x.type, x.value]));
  return { date: `${o.year}-${o.month}-${o.day}`, min: +o.hour * 60 + +o.minute };
}
const serialToIso = v => new Date(Math.round((v - 25569) * 864e5)).toISOString();
function toDate(v) {
  if (typeof v === "number") return serialToIso(v).slice(0, 10);
  const s = String(v).trim(), m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  return m ? `${m[3]}-${pad(m[2])}-${pad(m[1])}` : s.slice(0, 10);
}
function toHora(v) {
  if (v === "" || v == null) return "";
  if (typeof v === "number") { const m = Math.round((v % 1) * 1440); return `${pad(Math.floor(m / 60) % 24)}:${pad(m % 60)}`; }
  return String(v).trim().slice(0, 5);
}
const toReg = v => typeof v === "number" ? serialToIso(v).slice(0, 16).replace("T", " ") : String(v ?? "").trim();

// ---------- Normalización y deduplicado ----------
function norm(r) {
  const grupo = String(r.grupo).trim().toUpperCase(), nn = normName(r.sucursal);
  return {
    fecha: toDate(r.fecha), grupo, key: nn,
    sucursal: CANON_MAP.get(grupo + "|" + nn) || String(r.sucursal).trim().replace(/\s+/g, " "),
    estado: String(r.estado ?? "").trim(), hora: toHora(r.horaLlegada),
    punt: String(r.puntualidad ?? "").trim(), reg: toReg(r.registrado)
  };
}
// ok = a tiempo · late = tarde · no = no llegó
function st(r) {
  const p = normName(r.punt), e = normName(r.estado);
  if (p.includes("tiempo")) return "ok";
  if (p.includes("tarde") || e.includes("tarde")) return "late";
  if (p.includes("lleg") || e.startsWith("no")) return "no";
  return e.startsWith("env") ? "ok" : "no";
}
function dedupe(rows) { // una fila por (fecha, grupo, punto): prefiere Enviado, luego mayor Registrado
  const m = new Map();
  rows.map(norm).filter(r => /^\d{4}-\d{2}-\d{2}$/.test(r.fecha) && GRUPOS.includes(r.grupo)).forEach(r => {
    const k = `${r.fecha}|${r.grupo}|${r.key}`, p = m.get(k);
    if (!p) return m.set(k, r);
    const rs = st(r) !== "no", ps = st(p) !== "no";
    if (rs !== ps ? rs : r.reg >= p.reg) m.set(k, r);
  });
  return [...m.values()];
}

// ---------- Carga ----------
async function loadJson() {
  const r = await fetch(C.JSON_URL + "?t=" + Date.now());
  if (!r.ok) throw new Error("No se pudo leer " + C.JSON_URL);
  return r.json();
}
async function token() {
  const req = { scopes: ["Files.Read.All"], account: msal.getAllAccounts()[0] };
  try { return (await msal.acquireTokenSilent(req)).accessToken; }
  catch { return (await msal.acquireTokenPopup(req)).accessToken; }
}
async function loadGraph() {
  const t = await token();
  let url = `https://graph.microsoft.com/v1.0/users/${C.OWNER_UPN}/drive/items/${C.ITEM_ID}/workbook/tables/${C.TABLE}/rows?$top=1000`;
  const out = [];
  while (url) {
    const r = await fetch(url, { headers: { Authorization: "Bearer " + t } });
    if (!r.ok) throw new Error("Graph " + r.status + ": " + (await r.text()).slice(0, 200));
    const j = await r.json();
    j.value.forEach(x => { const v = x.values[0]; out.push({ fecha: v[0], grupo: v[1], sucursal: v[2], estado: v[3], registrado: v[4], horaLlegada: v[5], puntualidad: v[6] }); });
    url = j["@odata.nextLink"];
  }
  return out;
}
async function load() {
  try {
    if (C.SOURCE === "graph") {
      if (!msal.getAllAccounts().length) { $("login").hidden = false; show("Inicia sesión con tu cuenta Pompeyo para ver el dashboard."); return; }
      $("login").hidden = true; $("user").textContent = msal.getAllAccounts()[0].username;
    }
    DATA = dedupe(C.SOURCE === "graph" ? await loadGraph() : await loadJson());
    if (!DATA.length) { $("app").hidden = true; show("Aún no hay registros de stock."); return; }
    const last = DATA.reduce((a, r) => r.reg > a ? r.reg : a, "");
    $("upd").textContent = "Última actualización de Marco: " + (last ? ddmm(last) + last.slice(10) : "–") + " · página cargada " + new Date().toLocaleTimeString("es-CL", { timeZone: "America/Santiago", hour: "2-digit", minute: "2-digit" });
    $("msg").hidden = true; $("app").hidden = false;
    setupFilters(); render();
  } catch (e) { $("app").hidden = true; show("Error: " + e.message, true); }
}
function show(t, err) { $("msg").hidden = false; $("msg").className = err ? "err" : ""; $("msg").textContent = t; }

// ---------- Filtros ----------
const opt = (v, t) => `<option value="${v}">${t}</option>`;
function setupFilters() {
  const keep = { g: $("fGrupo").value, m: $("fMes").value, s: $("fSuc").value };
  $("fGrupo").innerHTML = opt("", "Todos") + GRUPOS.map(g => opt(g, g)).join("");
  const meses = [...new Set([...DATA.map(r => r.fecha.slice(0, 7)), nowSt().date.slice(0, 7)])].sort().reverse();
  $("fMes").innerHTML = meses.map(m => opt(m, MESES[+m.slice(5) - 1] + " " + m.slice(0, 4))).join("");
  if (keep.g) $("fGrupo").value = keep.g;
  if (keep.m && meses.includes(keep.m)) $("fMes").value = keep.m;
  fillSuc(keep.s);
}
function pointsOf(g) { // canónicos + cualquier nombre extra que aparezca en los datos
  const extra = [...new Set(DATA.filter(r => r.grupo === g).map(r => r.sucursal))].filter(n => !CANON[g].includes(n));
  return [...CANON[g], ...extra.sort()];
}
function fillSuc(sel) {
  const g = $("fGrupo").value, s = GRUPOS.filter(x => !g || x === g).flatMap(pointsOf);
  $("fSuc").innerHTML = opt("", "Todos") + s.map(x => opt(x, x)).join("");
  if (sel && s.includes(sel)) $("fSuc").value = sel;
}

// ---------- Render ----------
const ICON = { ok: "✓", late: "⏱", no: "✗" }, LABEL = { ok: "A tiempo", late: "Tarde", no: "No llegó" };
const CLS = { ok: "y", late: "l", no: "n" };
const isWE = iso => { const d = new Date(iso + "T12:00:00Z").getUTCDay(); return d === 0 || d === 6; };
const isProv = r => { const n = nowSt(); return st(r) === "no" && r.fecha === n.date && n.min >= CIERRE && n.min < BARRIDO; };

function render() {
  renderHoy();
  const mes = $("fMes").value, g = $("fGrupo").value, s = $("fSuc").value, we = $("fWE").checked;
  const inScope = r => (!g || r.grupo === g) && (!s || r.sucursal === s);
  const monthRows = DATA.filter(r => r.fecha.startsWith(mes) && inScope(r));
  const rows = monthRows.filter(r => we || !isWE(r.fecha)); // base de los indicadores

  // Indicadores
  const tot = rs => ({ n: rs.length, ok: rs.filter(r => st(r) === "ok").length, late: rs.filter(r => st(r) === "late").length });
  const kp = (t, label) => `<div class="card kpi"><span>${label}</span><b>${fmtPct(pct(t.ok, t.n))}</b><span>puntual · total ${fmtPct(pct(t.ok + t.late, t.n))}</span><span>${t.ok} a tiempo · <span class="lt">${t.late} tarde</span> · ${t.n - t.ok - t.late} no llegó</span><div class="bar"><i style="width:${pct(t.ok, t.n) || 0}%"></i></div></div>`;
  $("kpis").innerHTML = kp(tot(rows), "Cumplimiento puntual") + GRUPOS.filter(x => !g || x === g).map(x => kp(tot(rows.filter(r => r.grupo === x)), "Grupo " + x)).join("");

  // Matriz punto × día
  const [y, m] = mes.split("-").map(Number), dias = new Date(y, m, 0).getDate(), hoy = nowSt().date;
  const map = new Map(monthRows.map(r => [`${r.grupo}|${r.sucursal}|${+r.fecha.slice(8)}`, r]));
  let h = "<thead><tr><th>Punto</th>";
  for (let d = 1; d <= dias; d++) h += `<th class="${isWE(`${mes}-${pad(d)}`) ? "we" : ""}" title="${ddmm(`${mes}-${pad(d)}`)}">${d}</th>`;
  h += "<th title='% puntual'>%</th><th title='Días tarde'>⏱</th></tr></thead><tbody>";
  let any = false;
  GRUPOS.filter(x => !g || x === g).forEach(grp => {
    const pts = pointsOf(grp).filter(p => !s || p === s);
    if (!pts.length) return;
    h += `<tr class="g"><td colspan="${dias + 3}">${grp}</td></tr>`;
    pts.forEach(p => {
      any = true; h += `<tr><td><a href="#" class="pt" data-g="${grp}" data-p="${p}">${p}</a></td>`;
      let o = 0, n = 0, l = 0;
      for (let d = 1; d <= dias; d++) {
        const iso = `${mes}-${pad(d)}`, r = map.get(`${grp}|${p}|${d}`), cw = isWE(iso) ? " we" : "";
        if (!r) {
          const t = iso > hoy ? "Sin datos aún" : iso === hoy ? "Sin datos aún (cierre 12:00)" : "Sin dato (falló el cierre)";
          h += `<td class="c z${cw}" title="${t}">–</td>`; continue;
        }
        const e = st(r), prov = isProv(r);
        if (we || !isWE(iso)) { n++; if (e === "ok") o++; if (e === "late") l++; }
        const t = LABEL[e] + (r.hora ? " · " + r.hora : "") + (prov ? " · podría llegar hasta las 18:00" : "");
        h += `<td class="c ${CLS[e]}${prov ? " prov" : ""}${cw}" title="${t}">${ICON[e]}</td>`;
      }
      h += `<td class="pct">${fmtPct(pct(o, n))}</td><td class="lt">${l || ""}</td></tr>`;
    });
  });
  $("matrix").innerHTML = any ? h + "</tbody>" : "<tbody><tr><td>Sin datos para el filtro.</td></tr></tbody>";

  // Tendencia (puntual y total)
  const labels = [], vOk = [], vTot = [];
  for (let d = 1; d <= dias; d++) {
    const rs = rows.filter(r => +r.fecha.slice(8) === d), t = tot(rs);
    if (rs.length) { labels.push(`${pad(d)}-${pad(m)}`); vOk.push(pct(t.ok, t.n)); vTot.push(pct(t.ok + t.late, t.n)); }
  }
  if (chart) chart.destroy();
  chart = new Chart($("trend"), {
    type: "line",
    data: { labels, datasets: [
      { label: "Puntual", data: vOk, borderColor: "#1a7f4b", tension: .25, pointRadius: 3 },
      { label: "Total (incluye tarde)", data: vTot, borderColor: "#c98a00", borderDash: [5, 4], tension: .25, pointRadius: 2 }] },
    options: { scales: { y: { min: 0, max: 100, ticks: { callback: v => v + "%" } } }, plugins: { legend: { position: "bottom" } } }
  });

  // Ranking por cumplimiento puntual
  const rk = GRUPOS.filter(x => !g || x === g).flatMap(grp => pointsOf(grp).filter(p => !s || p === s).map(p => {
    const t = tot(rows.filter(r => r.grupo === grp && r.sucursal === p)); return { grp, p, ...t, v: pct(t.ok, t.n) };
  })).filter(r => r.n).sort((a, b) => b.v - a.v || b.n - a.n);
  $("rank").innerHTML = "<thead><tr><th>#</th><th>Punto</th><th>Grupo</th><th>% puntual</th></tr></thead><tbody>" +
    rk.map((r, i) => `<tr><td>${i + 1}</td><td>${r.p}</td><td>${r.grp}</td><td class="pct">${fmtPct(r.v)} <span class="sub">(${r.ok}/${r.n}${r.late ? " · " + r.late + " tarde" : ""})</span></td></tr>`).join("") + "</tbody>";
}

// Vista "Hoy": por grupo, cada punto con estado y hora
function renderHoy() {
  const n = nowSt(), dates = [...new Set(DATA.map(r => r.fecha))].sort();
  const hasToday = dates.includes(n.date), fecha = hasToday ? n.date : dates[dates.length - 1];
  const rows = DATA.filter(r => r.fecha === fecha);
  const last = rows.reduce((a, r) => r.reg > a ? r.reg : a, "");
  const title = hasToday ? "Hoy · " + ddmm(fecha) : "Último día con datos · " + ddmm(fecha);
  const note = hasToday ? (n.min < BARRIDO ? "Los \"No llegó\" de hoy aún pueden cambiar a \"Tarde\" hasta las 18:00 (barrido de rezagados)." : "Dato final.") : (n.min < CIERRE ? "Hoy: sin datos aún (el cierre es a las 12:00)." : "Hoy aún no tiene registros.");
  $("hoyT").innerHTML = `${title} <span class="sub">${note}${last ? " Actualizado: " + last.slice(11) : ""}</span>`;
  $("hoy").innerHTML = GRUPOS.map(grp => {
    const gr = rows.filter(r => r.grupo === grp), cnt = { ok: 0, late: 0, no: 0 };
    gr.forEach(r => cnt[st(r)]++);
    if (!gr.length) return `<div class="card"><h2>${grp}</h2><div class="sub">Sin dato</div></div>`;
    const lis = pointsOf(grp).map(p => {
      const r = gr.find(x => x.sucursal === p);
      if (!r) return `<li class="z"><span>–</span>${p}<em>sin dato</em></li>`;
      const e = st(r), prov = isProv(r);
      return `<li class="${CLS[e]}${prov ? " prov" : ""}"><span>${ICON[e]}</span>${p}<em>${r.hora ? r.hora + (e === "late" ? " · tarde" : "") : prov ? "podría llegar" : LABEL[e]}</em></li>`;
    }).join("");
    return `<div class="card"><h2>${grp} <span class="sub">${cnt.ok} a tiempo · ${cnt.late} tarde · ${cnt.no} no llegó</span></h2><ul class="hoyl">${lis}</ul></div>`;
  }).join("");
}

// Detalle de un punto
function detalle(grp, p) {
  const rs = DATA.filter(r => r.grupo === grp && r.sucursal === p).sort((a, b) => b.fecha.localeCompare(a.fecha));
  $("dT").textContent = `${p} (${grp})`;
  $("dB").innerHTML = "<thead><tr><th>Fecha</th><th>Estado</th><th>Hora llegada</th></tr></thead><tbody>" +
    rs.map(r => { const e = st(r); return `<tr><td>${ddmm(r.fecha)}</td><td class="${CLS[e]}">${ICON[e]} ${LABEL[e]}${isProv(r) ? " (podría llegar)" : ""}</td><td>${r.hora || "–"}</td></tr>`; }).join("") + "</tbody>";
  $("dlg").showModal();
}

// ---------- Init ----------
$("fGrupo").onchange = () => { fillSuc(); render(); };
$("fMes").onchange = $("fSuc").onchange = $("fWE").onchange = render;
$("refresh").onclick = load;
$("login").onclick = async () => { await msal.loginPopup({ scopes: ["Files.Read.All"] }); load(); };
$("matrix").onclick = e => { const a = e.target.closest("a.pt"); if (a) { e.preventDefault(); detalle(a.dataset.g, a.dataset.p); } };
$("dClose").onclick = () => $("dlg").close();

(async function init() {
  if (C.SOURCE === "graph") {
    msal = new msal.PublicClientApplication({
      auth: { clientId: C.CLIENT_ID, authority: "https://login.microsoftonline.com/" + C.TENANT_ID, redirectUri: C.REDIRECT_URI },
      cache: { cacheLocation: "sessionStorage" }
    });
    await msal.initialize();
  }
  load();
  // cada 5 min entre 12:00 y 18:30 (Santiago); fuera de ese rango, cada 30 min
  setInterval(() => { const m = nowSt().min; if ((m >= CIERRE && m <= BARRIDO + 30) || ++tick % 6 === 0) load(); }, 5 * 60000);
})();
