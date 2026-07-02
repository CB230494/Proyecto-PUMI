const ARCHIVOS_EXCEL = [
  "DR1 C AVANCE JUNIO.xlsx",
  "DR1 N AVANCE JUNIO.xlsx",
  "DR1 S AVANCE JUNIO.xlsx",
  "DR2 AVANCE JUNIO.xlsx",
  "DR3 AVANCE JUNIO.xlsx",
  "DR4 AVANCE JUNIO.xlsx",
  "DR5 AVANCE JUNIO.xlsx",
  "DR6 AVANCE JUNIO.xlsx",
  "DR7 AVANCE JUNIO.xlsx",
  "DR8 AVANCE JUNIO.xlsx",
  "DR9 AVANCE JUNIO.xlsx",
  "DR10 AVANCE JUNIO.xlsx",
  "DR11 AVANCE JUNIO.xlsx",
  "DR12 AVANCE JUNIO.xlsx"
];

const MAPA_REGION = {
  "DR1 C": "Dirección Regional 1 - San José Central",
  "DR1 N": "Dirección Regional 1 - San José Norte",
  "DR1 S": "Dirección Regional 1 - San José Sur",
  "DR2": "Dirección Regional 2 - Alajuela",
  "DR3": "Dirección Regional 3 - Cartago",
  "DR4": "Dirección Regional 4 - Heredia",
  "DR5": "Dirección Regional 5",
  "DR6": "Dirección Regional 6",
  "DR7": "Dirección Regional 7",
  "DR8": "Dirección Regional 8",
  "DR9": "Dirección Regional 9",
  "DR10": "Dirección Regional 10",
  "DR11": "Dirección Regional 11",
  "DR12": "Dirección Regional 12 - Caribe"
};

let datos = [];
let datosFiltrados = [];
let chartDelegaciones = null;
let chartRegiones = null;
let chartProgramas = null;

const $ = (id) => document.getElementById(id);

function normalizar(txt){
  return String(txt ?? "").trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/\s+/g," ");
}
function numero(valor){
  if(valor === null || valor === undefined || valor === "") return 0;
  if(typeof valor === "number") return isFinite(valor) ? valor : 0;
  const n = Number(String(valor).replace("%","").replace(",",".").trim());
  return isNaN(n) ? 0 : n;
}
function formato(n){ return Number(n || 0).toLocaleString("es-CR", {maximumFractionDigits:0}); }
function pct(n){ return `${(Number(n || 0)*100).toFixed(1)}%`; }
function estadoCumplimiento(p){
  if(p >= 0.50) return {clave:"verde", texto:"50% o más"};
  if(p >= 0.45) return {clave:"naranja", texto:"45% a menos de 50%"};
  return {clave:"rojo", texto:"Menor a 45%"};
}
function regionDesdeArchivo(nombre){
  const base = nombre.replace(/\.xlsx|\.xlsm|\.xls/gi, "").replace(/_/g," ").replace(/\s+/g," ").trim().toUpperCase();
  const claves = Object.keys(MAPA_REGION).sort((a,b)=>b.length-a.length);
  for(const clave of claves){
    if(base.startsWith(clave)) return MAPA_REGION[clave];
  }
  const m = base.match(/DR\s*\d+/);
  return m ? (MAPA_REGION[m[0].replace(/\s/g,"")] || m[0]) : base;
}
function ordenRegion(valor){
  const t = normalizar(valor);
  const n = (t.match(/REGIONAL\s+(\d+)/) || t.match(/DR\s*(\d+)/) || [null,999])[1];
  let sub = 0;
  if(Number(n) === 1){
    if(t.includes("CENTRAL")) sub = 0;
    else if(t.includes("NORTE")) sub = 1;
    else if(t.includes("SUR")) sub = 2;
    else sub = 9;
  }
  return [Number(n), sub, t];
}
function ordenDelegacion(valor){
  const t = normalizar(valor);
  const n = (t.match(/\bD\s*(\d+)/) || [null,9999])[1];
  return [Number(n), t];
}
function compararArray(a,b){
  for(let i=0;i<Math.max(a.length,b.length);i++){
    if(a[i] < b[i]) return -1;
    if(a[i] > b[i]) return 1;
  }
  return 0;
}
function colorEstado(clave){
  if(clave === "verde") return "#15803D";
  if(clave === "naranja") return "#F59E0B";
  return "#B42318";
}

async function cargarArchivoExcel(nombre){
  const url = encodeURI(nombre);
  const resp = await fetch(url, {cache:"no-store"});
  if(!resp.ok) throw new Error(`No se pudo cargar ${nombre}`);
  const buffer = await resp.arrayBuffer();
  const wb = XLSX.read(buffer, {type:"array"});
  const region = regionDesdeArchivo(nombre);
  const registros = [];

  for(const hoja of wb.SheetNames){
    if(normalizar(hoja).startsWith("TOTAL")) continue;
    const ws = wb.Sheets[hoja];
    const raw = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
    if(!raw || raw.length === 0) continue;

    let filaHeader = -1;
    for(let i=0; i<Math.min(12, raw.length); i++){
      const textoFila = raw[i].map(v=>normalizar(v)).join(" ");
      if(textoFila.includes("PROGRAMA") && textoFila.includes("META")) { filaHeader = i; break; }
    }

    let inicio = filaHeader >= 0 ? filaHeader + 1 : 2;
    let programaActual = "";

    for(let i=inicio; i<raw.length; i++){
      const row = raw[i];
      if(!row || row.every(v => String(v).trim()==="")) continue;
      const programaCelda = String(row[0] ?? "").trim();
      const actividad = String(row[1] ?? "").trim();
      if(programaCelda && normalizar(programaCelda) !== "NAN") programaActual = programaCelda;
      if(!programaActual || !actividad || ["TOTAL","TOTALES","NAN"].includes(normalizar(actividad))) continue;

      const meta = numero(row[2]);
      const avance = numero(row[3]);
      const pendiente = Math.max(meta - avance, 0);
      if(meta === 0 && avance === 0 && pendiente === 0) continue;
      const cumplimiento = meta ? avance / meta : 0;
      const estado = estadoCumplimiento(cumplimiento);

      registros.push({
        archivo:nombre,
        region,
        delegacion:hoja,
        programa:programaActual,
        actividad,
        meta,
        avance,
        pendiente,
        cumplimiento,
        estadoClave:estado.clave,
        estadoTexto:estado.texto
      });
    }
  }
  return registros;
}

async function cargarDatos(){
  $("estadoCarga").textContent = "Cargando archivos...";
  datos = [];
  const errores = [];
  for(const archivo of ARCHIVOS_EXCEL){
    try{
      const regs = await cargarArchivoExcel(archivo);
      datos.push(...regs);
    }catch(e){
      errores.push(archivo);
      console.warn(e);
    }
  }
  datos.sort((a,b)=> compararArray(ordenRegion(a.region), ordenRegion(b.region)) || compararArray(ordenDelegacion(a.delegacion), ordenDelegacion(b.delegacion)) || normalizar(a.programa).localeCompare(normalizar(b.programa)));
  $("estadoCarga").textContent = errores.length ? `Cargados ${ARCHIVOS_EXCEL.length-errores.length}/${ARCHIVOS_EXCEL.length}. Faltan: ${errores.length}` : `Cargados ${ARCHIVOS_EXCEL.length} Excel regionales`;
  poblarFiltros();
  aplicarFiltros();
}

function opcionesUnicas(campo, base=datos){
  const vals = [...new Set(base.map(d=>d[campo]).filter(Boolean))];
  if(campo === "region") return vals.sort((a,b)=>compararArray(ordenRegion(a), ordenRegion(b)));
  if(campo === "delegacion") return vals.sort((a,b)=>compararArray(ordenDelegacion(a), ordenDelegacion(b)));
  return vals.sort((a,b)=>normalizar(a).localeCompare(normalizar(b)));
}
function llenarSelect(id, opciones, inicial){
  const sel = $(id);
  const actual = sel.value;
  sel.innerHTML = `<option value="">${inicial}</option>` + opciones.map(o=>`<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`).join("");
  if(opciones.includes(actual)) sel.value = actual;
}
function poblarFiltros(){
  llenarSelect("filtroRegion", opcionesUnicas("region"), "Todas");
  llenarSelect("filtroDelegacion", opcionesUnicas("delegacion"), "Todas");
  llenarSelect("filtroPrograma", opcionesUnicas("programa"), "Todos");
}
function aplicarFiltros(){
  const region = $("filtroRegion").value;
  const delegacion = $("filtroDelegacion").value;
  const programa = $("filtroPrograma").value;
  const estado = $("filtroEstado").value;

  let baseDeleg = datos;
  if(region) baseDeleg = baseDeleg.filter(d=>d.region === region);
  llenarSelect("filtroDelegacion", opcionesUnicas("delegacion", baseDeleg), "Todas");
  if(delegacion && !opcionesUnicas("delegacion", baseDeleg).includes(delegacion)) $("filtroDelegacion").value = "";

  const delegacionFinal = $("filtroDelegacion").value;
  datosFiltrados = datos.filter(d =>
    (!region || d.region === region) &&
    (!delegacionFinal || d.delegacion === delegacionFinal) &&
    (!programa || d.programa === programa) &&
    (!estado || d.estadoClave === estado)
  );
  actualizarDashboard();
  actualizarTabla();
  actualizarResumenInforme();
}

function agrupar(base, campo){
  const map = new Map();
  for(const d of base){
    const k = d[campo] || "Sin dato";
    if(!map.has(k)) map.set(k,{nombre:k, meta:0, avance:0, pendiente:0, actividades:0});
    const x = map.get(k);
    x.meta += d.meta; x.avance += d.avance; x.pendiente += d.pendiente; x.actividades += 1;
  }
  return [...map.values()].map(x=>({...x, cumplimiento:x.meta ? x.avance/x.meta : 0, estadoClave:estadoCumplimiento(x.meta ? x.avance/x.meta : 0).clave}));
}
function actualizarDashboard(){
  const meta = datosFiltrados.reduce((s,d)=>s+d.meta,0);
  const avance = datosFiltrados.reduce((s,d)=>s+d.avance,0);
  const pendiente = Math.max(meta-avance,0);
  const cumplimiento = meta ? avance/meta : 0;
  $("mMeta").textContent = formato(meta);
  $("mAvance").textContent = formato(avance);
  $("mPendiente").textContent = formato(pendiente);
  $("mPorcentaje").textContent = pct(cumplimiento);
  $("mRegiones").textContent = new Set(datosFiltrados.map(d=>d.region)).size;
  $("mDelegaciones").textContent = new Set(datosFiltrados.map(d=>d.delegacion)).size;

  const porDelegacion = agrupar(datosFiltrados,"delegacion").sort((a,b)=>b.cumplimiento-a.cumplimiento);
  actualizarTop(porDelegacion);
  crearGraficoDelegaciones(porDelegacion);
  crearGraficoSimple("graficoRegiones", agrupar(datosFiltrados,"region").sort((a,b)=>compararArray(ordenRegion(a.nombre), ordenRegion(b.nombre))), "Cumplimiento por región");
  crearGraficoSimple("graficoProgramas", agrupar(datosFiltrados,"programa").sort((a,b)=>b.cumplimiento-a.cumplimiento), "Cumplimiento por programa");
}
function actualizarTop(lista){
  const mejores = lista.filter(x=>x.meta>0).slice(0,2);
  const peores = lista.filter(x=>x.meta>0).sort((a,b)=>a.cumplimiento-b.cumplimiento).slice(0,2);
  $("topMejores").innerHTML = mejores.map(x=>`<li>${escapeHtml(x.nombre)} — ${pct(x.cumplimiento)}</li>`).join("") || "<li>Sin datos</li>";
  $("topPeores").innerHTML = peores.map(x=>`<li>${escapeHtml(x.nombre)} — ${pct(x.cumplimiento)}</li>`).join("") || "<li>Sin datos</li>";
}
function crearGraficoDelegaciones(lista){
  const labels = lista.map(x=>x.nombre);
  const meta = lista.map(x=>x.meta);
  const avance = lista.map(x=>x.avance);
  const pendiente = lista.map(x=>x.pendiente);
  const linea50 = lista.map(x=>x.meta*0.5);
  if(chartDelegaciones) chartDelegaciones.destroy();
  chartDelegaciones = new Chart($("graficoDelegaciones"), {
    type:"bar",
    data:{labels, datasets:[
      {label:"Meta", data:meta, backgroundColor:"rgba(255,255,255,.9)", borderColor:"#002B7F", borderWidth:2},
      {label:"Avance", data:avance, backgroundColor:"#002B7F"},
      {label:"Pendiente", data:pendiente, backgroundColor:"#B42318"},
      {label:"Límite 50%", data:linea50, type:"line", borderColor:"#F59E0B", borderWidth:3, borderDash:[8,6], pointRadius:0}
    ]},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:"top"}}, scales:{x:{ticks:{maxRotation:70,minRotation:45}}, y:{beginAtZero:true}}}
  });
}
function crearGraficoSimple(canvasId, lista, titulo){
  const chartVar = canvasId === "graficoRegiones" ? "chartRegiones" : "chartProgramas";
  if(window[chartVar]) window[chartVar].destroy();
  const labels = lista.map(x=>x.nombre);
  const porcentajes = lista.map(x=>Number((x.cumplimiento*100).toFixed(1)));
  const colores = lista.map(x=>colorEstado(x.estadoClave));
  window[chartVar] = new Chart($(canvasId), {
    type:"bar",
    data:{labels, datasets:[{label:"% cumplimiento", data:porcentajes, backgroundColor:colores}]},
    options:{responsive:true, plugins:{legend:{display:false}, title:{display:false,text:titulo}}, scales:{y:{beginAtZero:true, max:100, ticks:{callback:v=>v+"%"}}, x:{ticks:{maxRotation:45,minRotation:20}}}}
  });
  if(canvasId === "graficoRegiones") chartRegiones = window[chartVar]; else chartProgramas = window[chartVar];
}
function actualizarTabla(){
  const tbody = $("tablaDatos").querySelector("tbody");
  tbody.innerHTML = datosFiltrados.map(d=>`
    <tr class="estado-${d.estadoClave}">
      <td>${escapeHtml(d.region)}</td>
      <td>${escapeHtml(d.delegacion)}</td>
      <td>${escapeHtml(d.programa)}</td>
      <td>${escapeHtml(d.actividad)}</td>
      <td>${formato(d.meta)}</td>
      <td>${formato(d.avance)}</td>
      <td>${formato(d.pendiente)}</td>
      <td>${pct(d.cumplimiento)}</td>
      <td><span class="badge ${d.estadoClave}">${escapeHtml(d.estadoTexto)}</span></td>
    </tr>`).join("");
}
function actualizarResumenInforme(){
  const meta = datosFiltrados.reduce((s,d)=>s+d.meta,0);
  const avance = datosFiltrados.reduce((s,d)=>s+d.avance,0);
  const pendiente = Math.max(meta-avance,0);
  $("resumenInforme").innerHTML = `
    <div class="preview-card"><strong>Registros incluidos:</strong> ${datosFiltrados.length}</div>
    <div class="preview-card"><strong>Meta:</strong> ${formato(meta)} | <strong>Avance:</strong> ${formato(avance)} | <strong>Pendiente:</strong> ${formato(pendiente)} | <strong>Cumplimiento:</strong> ${pct(meta?avance/meta:0)}</div>
  `;
}

function generarPdf(){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({orientation:"landscape", unit:"pt", format:"letter"});
  const meta = datosFiltrados.reduce((s,d)=>s+d.meta,0);
  const avance = datosFiltrados.reduce((s,d)=>s+d.avance,0);
  const pendiente = Math.max(meta-avance,0);
  const cumplimiento = meta ? avance/meta : 0;
  const filtros = [
    ["Dirección Regional", $("filtroRegion").value || "Todas"],
    ["Delegación", $("filtroDelegacion").value || "Todas"],
    ["Programa", $("filtroPrograma").value || "Todos"],
    ["Cumplimiento", $("filtroEstado").selectedOptions[0].text || "Todos"]
  ];

  doc.setFillColor(0,43,127); doc.rect(0,0,792,58,"F");
  doc.setTextColor(255,255,255); doc.setFont("helvetica","bold"); doc.setFontSize(18);
  doc.text("Informe Nacional PUMI 2026", 396, 36, {align:"center"});
  doc.setTextColor(0,43,127); doc.setFontSize(22); doc.text("Dashboard nacional de metas y avances", 396, 105, {align:"center"});
  doc.setFontSize(11); doc.setFont("helvetica","normal"); doc.setTextColor(40,40,40);
  doc.text("Informe generado con los datos filtrados actualmente en la aplicación.", 396, 128, {align:"center"});

  doc.autoTable({startY:155, head:[["Indicador","Resultado"]], body:[
    ["Meta oficial", formato(meta)], ["Avance", formato(avance)], ["Pendiente", formato(pendiente)], ["% cumplimiento", pct(cumplimiento)], ["Registros incluidos", datosFiltrados.length]
  ], styles:{fontSize:9}, headStyles:{fillColor:[0,43,127], textColor:[255,255,255]}, alternateRowStyles:{fillColor:[244,247,251]}, margin:{left:70,right:70}});

  doc.autoTable({startY:doc.lastAutoTable.finalY+18, head:[["Filtro","Valor aplicado"]], body:filtros, styles:{fontSize:9}, headStyles:{fillColor:[184,138,42], textColor:[255,255,255]}, margin:{left:70,right:70}});

  doc.addPage("landscape");
  doc.setFillColor(0,43,127); doc.rect(0,0,792,44,"F");
  doc.setTextColor(255,255,255); doc.setFont("helvetica","bold"); doc.setFontSize(14); doc.text("Detalle filtrado con semáforo de cumplimiento", 396, 28, {align:"center"});

  const body = datosFiltrados.slice(0, 300).map(d=>[d.region,d.delegacion,d.programa,d.actividad,formato(d.meta),formato(d.avance),formato(d.pendiente),pct(d.cumplimiento),d.estadoTexto]);
  doc.autoTable({
    startY:60,
    head:[["Región","Delegación","Programa","Actividad","Meta","Avance","Pendiente","%","Estado"]],
    body,
    styles:{fontSize:6, cellPadding:3},
    headStyles:{fillColor:[0,43,127], textColor:[255,255,255]},
    didParseCell:function(data){
      if(data.section === "body"){
        const estado = datosFiltrados[data.row.index]?.estadoClave;
        if(estado === "verde") data.cell.styles.fillColor = [236,253,245];
        if(estado === "naranja") data.cell.styles.fillColor = [255,247,237];
        if(estado === "rojo") data.cell.styles.fillColor = [254,242,242];
      }
    },
    margin:{left:25,right:25}
  });

  const nombre = `INFORME_NACIONAL_PUMI_2026_${new Date().toISOString().slice(0,10)}.pdf`;
  doc.save(nombre);
}

function exportCsv(){
  const header = ["Region","Delegacion","Programa","Actividad","Meta","Avance","Pendiente","Cumplimiento","Estado"];
  const rows = datosFiltrados.map(d=>[d.region,d.delegacion,d.programa,d.actividad,d.meta,d.avance,d.pendiente,(d.cumplimiento*100).toFixed(1)+"%",d.estadoTexto]);
  const csv = [header, ...rows].map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob(["\ufeff"+csv], {type:"text/csv;charset=utf-8"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = "DASHBOARD_NACIONAL_PUMI_2026.csv"; a.click();
  URL.revokeObjectURL(a.href);
}
function escapeHtml(s){
  return String(s ?? "").replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
}

function activarSeccion(id){
  document.querySelectorAll(".section").forEach(s=>s.classList.remove("active"));
  document.querySelectorAll(".menu-btn").forEach(b=>b.classList.remove("active"));
  $(id).classList.add("active");
  document.querySelector(`.menu-btn[data-section="${id}"]`)?.classList.add("active");
}

document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll(".menu-btn").forEach(btn=>btn.addEventListener("click",()=>activarSeccion(btn.dataset.section)));
  document.querySelectorAll(".quick-card").forEach(btn=>btn.addEventListener("click",()=>activarSeccion(btn.dataset.go)));
  ["filtroRegion","filtroDelegacion","filtroPrograma","filtroEstado"].forEach(id=>$(id).addEventListener("change",aplicarFiltros));
  $("btnRecargar").addEventListener("click", cargarDatos);
  $("btnGenerarPdf").addEventListener("click", generarPdf);
  $("btnExportCsv").addEventListener("click", exportCsv);
  cargarDatos();
});
