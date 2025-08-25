// Ruta del Excel local
const EXCEL_PATH = "./Base de datos Grupo PRA.xlsx";

// UI
const PANEL           = document.getElementById("panel");
const FILTERS_WRAP    = document.getElementById("filtros");
const CHIPS           = document.getElementById("chips");
const TOGGLE_PREVIEW  = document.getElementById("toggle-preview");
const PREVIEW         = document.getElementById("preview");
const PREVIEW_CONTENT = document.getElementById("preview-content");

// Estado
let ALL_ROWS = [];
let KEYS = { grupoKey:null, subsectorKey:null };
let subsectorChart = null;

/* Utilitarios UI */
function showMessage(msg){ PANEL.innerHTML = `<p class="muted">${msg}</p>`; }
function countsByKey(rows, key){
  const map = new Map();
  rows.forEach(r=>{
    let v = r[key];
    if (v == null) return;
    v = String(v).trim();
    if (!v) return;
    map.set(v, (map.get(v) || 0) + 1);
  });
  return Array.from(map.entries()).sort((a,b)=>b[1]-a[1]);
}
function buildTable(rows){
  if(!rows.length) return document.createTextNode("No hay filas.");
  const headers = Object.keys(rows[0]);
  const table = document.createElement("table");
  table.className = "table";
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  headers.forEach(h=>{ const th=document.createElement("th"); th.textContent=h; trh.appendChild(th); });
  thead.appendChild(trh);
  const tbody = document.createElement("tbody");
  rows.slice(0,200).forEach(r=>{
    const tr=document.createElement("tr");
    headers.forEach(h=>{
      const td=document.createElement("td");
      td.textContent = r[h]==null ? "" : String(r[h]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(thead); table.appendChild(tbody);
  return table;
}

/* Normalización robusta */
function norm(s){
  return String(s ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .toLowerCase().replace(/\s+/g," ").trim()
    .replace(/[^a-z0-9 ]/g,"");
}
function findColFlexible(headers, patterns){
  const normHeaders = headers.map(h => ({raw:h, n:norm(h)}));
  for(const pat of patterns){
    const np = norm(pat);
    const exact = normHeaders.find(h => h.n === np);
    if(exact) return exact.raw;
    const incl = normHeaders.find(h => h.n.includes(np));
    if(incl) return incl.raw;
  }
  return null;
}

/* Leer Excel */
async function readLocalXlsx(){
  const res = await fetch(EXCEL_PATH);
  if(!res.ok) throw new Error("No se pudo cargar el XLSX. ¿Abriste con Live Server?");
  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, { type:"array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: null });
}

/* Chips */
function renderGroupChips(groups){
  CHIPS.innerHTML = "";
  const allChip = document.createElement("button");
  allChip.className = "chip is-active";
  allChip.textContent = "Todos";
  allChip.addEventListener("click", ()=>applyGroupFilter("Todos"));
  CHIPS.appendChild(allChip);

  groups.forEach(g=>{
    const b = document.createElement("button");
    b.className = "chip";
    b.textContent = g;
    b.title = `grupo = ${g}`;
    b.addEventListener("click", ()=>applyGroupFilter(g));
    CHIPS.appendChild(b);
  });

  FILTERS_WRAP.hidden = false;
}
function setActiveChip(value){
  [...CHIPS.querySelectorAll(".chip")].forEach(ch=>{
    ch.classList.toggle("is-active", ch.textContent === value);
  });
}

/* Filtro y render */
function applyGroupFilter(value){
  setActiveChip(value);
  const filtered = value==="Todos" || !KEYS.grupoKey
    ? ALL_ROWS
    : ALL_ROWS.filter(r => String(r[KEYS.grupoKey] ?? "").trim() === value);

  renderSubsectorChart(filtered);
  PREVIEW_CONTENT.innerHTML = "";
  PREVIEW_CONTENT.appendChild(buildTable(filtered));
}

/* Gráfico pastel */
function renderSubsectorChart(rows){
  if(!rows.length || !KEYS.subsectorKey) return;
  const data = countsByKey(rows, KEYS.subsectorKey);
  const ctx = document.getElementById("subsectorChart").getContext("2d");
  if(subsectorChart) subsectorChart.destroy();
  subsectorChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: data.map(d=>d[0]),
      datasets: [{
        data: data.map(d=>d[1]),
        backgroundColor: ["#3b82f6","#f43f5e","#10b981","#f59e0b","#6366f1","#a78bfa","#ef4444","#22c55e"]
      }]
    },
    options: {
      responsive:true,
      maintainAspectRatio:true,   // respeta 280x280 del CSS
      plugins:{ legend:{ position:"bottom" } },
      cutout:"55%"
    }
  });
}

/* Toggle preview */
TOGGLE_PREVIEW?.addEventListener("click", ()=>{
  const collapsed = PREVIEW.classList.toggle("collapsed");
  PREVIEW.setAttribute("aria-expanded", String(!collapsed));
});

/* Inicio: cargar Excel y preparar UI */
(async function init(){
  showMessage("Cargando Excel local…");
  try{
    ALL_ROWS = await readLocalXlsx();
    if(!ALL_ROWS.length){ showMessage("El Excel está vacío."); return; }

    const headers = Object.keys(ALL_ROWS[0]);
    KEYS.grupoKey     = findColFlexible(headers, ["grupo","grupos"]);
    KEYS.subsectorKey = findColFlexible(headers, ["subsector","sub sector","sector"]);

    if(!KEYS.subsectorKey){
      showMessage("No se encontró la columna ‘subsector’ en el Excel.");
      return;
    }

    // chips
    const groups = new Set();
    if(KEYS.grupoKey){
      ALL_ROWS.forEach(r=>{
        const v = String(r[KEYS.grupoKey] ?? "").trim();
        if(v) groups.add(v);
      });
    }
    renderGroupChips([...groups].sort((a,b)=>a.localeCompare(b,"es")));

    // render inicial
    applyGroupFilter("Todos");
    showMessage("");
  }catch(e){
    console.error(e);
    showMessage("No fue posible leer el Excel. Abre con Live Server.");
  }
})();








