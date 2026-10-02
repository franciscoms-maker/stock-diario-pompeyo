const C = window.CONFIG;
const $ = id => document.getElementById(id);
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const GRUPOS = ["KIA", "SDA", "ADP"];
let DATA = [], chart = null, msal = null;

// ---------- Carga de datos ----------
function excelDate(v) {
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  return m ? `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` : s.slice(0, 10);
}
const norm = r => ({
  fecha: excelDate(r.fecha), grupo: String(r.grupo).trim().toUpperCase(),
  sucursal: String(r.sucursal).trim(), estado: String(r.estado).trim(), reg: r.reg ?? ""
});

function dedupe(rows) { // una fila por (grupo, sucursal, fecha); gana el último registro
  const m = new Map();
  rows.map(norm).filter(r => /^\d{4}-\d{2}-\d{2}$/.test(r.fecha)).forEach((r, i) => {
    const k = `${r.grupo}|${r.sucursal}|${r.fecha}`, p = m.get(k);
    if (!p || String(r.reg) >= String(p.reg) || (r.reg === "" && p.reg === "")) m.set(k, r);
  });
  return [...m.values()];
}

async function loadJson() {
  const r = await fetch(C.JSON_URL + "?t=" + Date.now());
  if (!r.ok) throw new Error("No se pudo leer " + C.JSON_URL);
  return (await r.json()).map(x => ({ ...x, reg: x.registrado }));
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
    j.value.forEach(x => { const v = x.values[0]; out.push({ fecha: v[0], grupo: v[1], sucursal: v[2], estado: v[3], reg: v[4] }); });
    url = j["@odata.nextLink"];
  }
  return out;
}

async function load() {
  try {
    $("msg").hidden = false; $("msg").textContent = "Cargando datos…"; $("msg").className = "";
    if (C.SOURCE === "graph") {
      if (!msal.getAllAccounts().length) { $("login").hidden = false; $("msg").textContent = "Inicia sesión con tu cuenta Pompeyo para ver el dashboard."; return; }
      $("login").hidden = true; $("user").textContent = msal.getAllAccounts()[0].username;
    }
    DATA = dedupe(C.SOURCE === "graph" ? await loadGraph() : await loadJson());
    $("upd").textContent = "Actualizado " + new Date().toLocaleString("es-CL", { timeZone: "America/Santiago" });
    $("msg").hidden = true; $("app").hidden = false;
    setupFilters(); render();
  } catch (e) { $("msg").hidden = false; $("msg").className = "err"; $("msg").textContent = "Error: " + e.message; }
}

// ---------- Filtros ----------
const opt = (v, t) => `<option value="${v}">${t}</option>`;
function setupFilters() {
  const keep = { g: $("fGrupo").value, m: $("fMes").value, s: $("fSuc").value };
  $("fGrupo").innerHTML = opt("", "Todos") + GRUPOS.map(g => opt(g, g)).join("");
  const meses = [...new Set(DATA.map(r => r.fecha.slice(0, 7)))].sort().reverse();
  $("fMes").innerHTML = meses.map(m => opt(m, MESES[+m.slice(5) - 1] + " " + m.slice(0, 4))).join("");
  if (keep.g) $("fGrupo").value = keep.g;
  if (keep.m && meses.includes(keep.m)) $("fMes").value = keep.m;
  fillSuc(keep.s);
}
function fillSuc(sel) {
  const g = $("fGrupo").value;
  const s = [...new Set(DATA.filter(r => !g || r.grupo === g).map(r => r.sucursal))].sort();
  $("fSuc").innerHTML = opt("", "Todas") + s.map(x => opt(x, x)).join("");
  if (sel && s.includes(sel)) $("fSuc").value = sel;
}

// ---------- Render ----------
const pct = (ok, tot) => tot ? Math.round(ok / tot * 100) : null;
const fmtPct = p => p === null ? "–" : p + "%";
const ddmm = iso => iso.slice(8) + "-" + iso.slice(5, 7) + "-" + iso.slice(0, 4);

function render() {
  const mes = $("fMes").value, g = $("fGrupo").value, s = $("fSuc").value;
  const rows = DATA.filter(r => r.fecha.startsWith(mes) && (!g || r.grupo === g) && (!s || r.sucursal === s));
  const ok = r => r.estado.toLowerCase() === "enviado";

  // KPIs
  const tot = (rs) => ({ ok: rs.filter(ok).length, n: rs.length });
  const kp = (t, label) => `<div class="card kpi"><span>${label}</span><b>${fmtPct(pct(t.ok, t.n))}</b><span>${t.ok} de ${t.n} registros</span><div class="bar"><i style="width:${pct(t.ok, t.n) || 0}%"></i></div></div>`;
  $("kpis").innerHTML = kp(tot(rows), "Cumplimiento total") +
    GRUPOS.filter(x => !g || x === g).map(x => kp(tot(rows.filter(r => r.grupo === x)), "Grupo " + x)).join("");

  // Matriz
  const [y, m] = mes.split("-").map(Number), dias = new Date(y, m, 0).getDate();
  const map = new Map(rows.map(r => [`${r.grupo}|${r.sucursal}|${+r.fecha.slice(8)}`, r]));
  const sucs = [...new Map(rows.map(r => [r.grupo + "|" + r.sucursal, r])).values()]
    .sort((a, b) => GRUPOS.indexOf(a.grupo) - GRUPOS.indexOf(b.grupo) || a.sucursal.localeCompare(b.sucursal));
  let h = "<thead><tr><th>Sucursal</th>";
  for (let d = 1; d <= dias; d++) h += `<th title="${ddmm(`${mes}-${String(d).padStart(2, "0")}`)}">${d}</th>`;
  h += "<th>%</th></tr></thead><tbody>";
  let cur = "";
  sucs.forEach(sc => {
    if (sc.grupo !== cur) { cur = sc.grupo; h += `<tr class="g"><td colspan="${dias + 2}">${cur}</td></tr>`; }
    h += `<tr><td>${sc.sucursal}</td>`; let o = 0, n = 0;
    for (let d = 1; d <= dias; d++) {
      const r = map.get(`${sc.grupo}|${sc.sucursal}|${d}`);
      const we = [0, 6].includes(new Date(y, m - 1, d).getDay()) ? " we" : "";
      if (!r) h += `<td class="c${we}"></td>`;
      else { n++; if (ok(r)) { o++; h += `<td class="c y${we}" title="Enviado">✓</td>`; } else h += `<td class="c n${we}" title="No enviado">✗</td>`; }
    }
    h += `<td class="pct">${fmtPct(pct(o, n))}</td></tr>`;
  });
  $("matrix").innerHTML = sucs.length ? h + "</tbody>" : "<tbody><tr><td>Sin datos para el filtro.</td></tr></tbody>";

  // Tendencia
  const labels = [], vals = [];
  for (let d = 1; d <= dias; d++) {
    const rs = rows.filter(r => +r.fecha.slice(8) === d);
    if (rs.length) { labels.push(String(d).padStart(2, "0") + "-" + String(m).padStart(2, "0")); vals.push(pct(rs.filter(ok).length, rs.length)); }
  }
  if (chart) chart.destroy();
  chart = new Chart($("trend"), {
    type: "line",
    data: { labels, datasets: [{ label: "% cumplimiento", data: vals, borderColor: "#1f4fd8", backgroundColor: "#1f4fd822", fill: true, tension: .25, pointRadius: 3 }] },
    options: { scales: { y: { min: 0, max: 100, ticks: { callback: v => v + "%" } } }, plugins: { legend: { display: false } } }
  });

  // Ranking
  const rk = sucs.map(sc => { const rs = rows.filter(r => r.grupo === sc.grupo && r.sucursal === sc.sucursal); const t = tot(rs); return { ...sc, p: pct(t.ok, t.n), ok: t.ok, n: t.n }; })
    .sort((a, b) => (b.p ?? -1) - (a.p ?? -1));
  $("rank").innerHTML = "<thead><tr><th>#</th><th>Sucursal</th><th>Grupo</th><th>%</th></tr></thead><tbody>" +
    rk.map((r, i) => `<tr><td>${i + 1}</td><td>${r.sucursal}</td><td>${r.grupo}</td><td class="pct">${fmtPct(r.p)} <span class="sub">(${r.ok}/${r.n})</span></td></tr>`).join("") + "</tbody>";
}

// ---------- Init ----------
$("fGrupo").onchange = () => { fillSuc(); render(); };
$("fMes").onchange = $("fSuc").onchange = render;
$("refresh").onclick = load;
$("login").onclick = async () => { await msal.loginPopup({ scopes: ["Files.Read.All"] }); load(); };

(async function init() {
  if (C.SOURCE === "graph") {
    msal = new msal.PublicClientApplication({
      auth: { clientId: C.CLIENT_ID, authority: "https://login.microsoftonline.com/" + C.TENANT_ID, redirectUri: C.REDIRECT_URI },
      cache: { cacheLocation: "sessionStorage" }
    });
    await msal.initialize();
  }
  load();
  setInterval(load, C.REFRESH_MIN * 60000);
})();
