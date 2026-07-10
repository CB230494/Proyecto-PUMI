import { CONFIG } from "./config.js";
import { ArcGISService, startOAuth, consumeOAuthHash } from "./arcgis.js";

const svc = new ArcGISService();
const state = {
  portalUser: null,
  user: null,
  activities: [],
  summary: [],
  catalogs: [],
  notifications: [],
  selectedActivity: null,
  demo: false,
  map: null,
  nationalMap: null
};

const $ = id => document.getElementById(id);
const attr = (obj, key, fallback = "") => obj?.[CONFIG.fieldMap[key]] ?? fallback;
const normalize = value => String(value ?? "").trim().toLowerCase();
const escapeHtml = s => String(s ?? "").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));

document.addEventListener("DOMContentLoaded", init);

async function init() {
  bindEvents();
  const token = consumeOAuthHash();
  if (token) svc.setToken(token);

  if (svc.token) {
    try {
      await loginFromToken();
      return;
    } catch (error) {
      svc.setToken("");
      showToast(error.message, true);
    }
  }
  showLogin();
}

function bindEvents() {
  $("btnLoginArcGIS").addEventListener("click", () => {
    if (CONFIG.oauthClientId.includes("COLOQUE")) {
      showToast("Primero configure el Client ID en js/config.js.", true);
      return;
    }
    startOAuth();
  });
  $("btnDemo").addEventListener("click", loginDemo);
  $("btnLogout").addEventListener("click", logout);
  $("btnRefresh").addEventListener("click", loadAll);
  $("activityForm").addEventListener("submit", submitActivity);
  $("userForm").addEventListener("submit", submitUser);
  $("btnLocation").addEventListener("click", getLocation);
  $("btnNotifications").addEventListener("click", () => $("notificationPanel").classList.remove("hidden"));
  $("btnCloseNotifications").addEventListener("click", () => $("notificationPanel").classList.add("hidden"));
  $("btnCloseModal").addEventListener("click", closeReviewModal);
  $("btnCancelReview").addEventListener("click", closeReviewModal);
  $("reviewForm").addEventListener("submit", submitReview);
  $("btnApplyFilters").addEventListener("click", renderReviewTable);
  $("programSelect").addEventListener("change", populateActivityOptions);
  $("provinceSelect").addEventListener("change", populateCantons);
  $("cantonSelect").addEventListener("change", populateDistricts);
}

async function loginFromToken() {
  state.portalUser = await svc.getPortalUser();
  state.user = await svc.getAppUser(state.portalUser.username);
  if (!state.user) throw new Error("Su usuario ArcGIS no está autorizado en PUMI_USUARIOS.");
  showMain();
  await loadAll();
}

function loginDemo() {
  state.demo = true;
  state.portalUser = { username: "nacional_demo", fullName: "Usuario Nacional Demo" };
  state.user = {
    usuario: "nacional_demo",
    nombre: "Usuario Nacional Demo",
    rol: "Administrador",
    direccion_regional: "",
    delegacion: "",
    programa: "",
    activo: 1
  };
  seedDemo();
  showMain();
  renderAll();
}

function showLogin() {
  $("loginView").classList.remove("hidden");
  $("mainView").classList.add("hidden");
}

function showMain() {
  $("loginView").classList.add("hidden");
  $("mainView").classList.remove("hidden");
  $("userName").textContent = attr(state.user, "nombre", state.portalUser?.fullName || "Usuario");
  $("userRole").textContent = attr(state.user, "rol", "");
  $("currentScope").textContent = [attr(state.user, "direccionRegional"), attr(state.user, "delegacion"), attr(state.user, "programa")].filter(Boolean).join(" · ");
  buildNavigation();
}

function logout() {
  svc.setToken("");
  sessionStorage.clear();
  location.reload();
}

function buildNavigation() {
  const role = normalize(attr(state.user, "rol"));
  const items = [{ id: "dashboardSection", label: "📊 Panel principal" }];

  if (role.includes("deleg")) items.push({ id: "registerSection", label: "➕ Registrar actividad" });
  if (role.includes("regional") || role.includes("coordin") || role.includes("nacional") || role.includes("admin")) {
    items.push({ id: "reviewSection", label: "✅ Revisión y validación" });
  }
  if (role.includes("nacional") || role.includes("admin")) items.push({ id: "nationalSection", label: "🗺️ Vista nacional" });
  if (role.includes("admin")) items.push({ id: "usersSection", label: "👥 Usuarios" });

  $("mainNav").innerHTML = items.map((item, i) =>
    `<button class="nav-btn ${i===0?"active":""}" data-section="${item.id}">${item.label}</button>`
  ).join("");

  document.querySelectorAll(".nav-btn").forEach(btn => btn.addEventListener("click", () => {
    document.querySelectorAll(".nav-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    document.querySelectorAll(".section").forEach(s => s.classList.remove("active"));
    $(btn.dataset.section).classList.add("active");
    if (btn.dataset.section === "nationalSection") renderNational();
  }));
}

async function loadAll() {
  try {
    const [acts, summary, catalogs] = await Promise.all([
      svc.query(CONFIG.layers.actividades),
      svc.query(CONFIG.layers.resumen),
      svc.query(CONFIG.layers.catalogos, { returnGeometry: false })
    ]);
    state.activities = acts;
    state.summary = summary;
    state.catalogs = catalogs.map(f => f.attributes);
    await loadNotifications();
    renderAll();
    showToast("Datos actualizados.");
  } catch (error) {
    showToast(error.message, true);
  }
}

async function loadNotifications() {
  const role = attr(state.user, "rol");
  const user = attr(state.user, "usuario");
  state.notifications = buildDerivedNotifications();
  if (CONFIG.layers.notificaciones && !state.demo) {
    try {
      const rows = await svc.query(CONFIG.layers.notificaciones, {
        where: `usuario_destino='${String(user).replaceAll("'","''")}' OR rol_destino='${String(role).replaceAll("'","''")}'`,
        returnGeometry: false
      });
      state.notifications.push(...rows.map(f => f.attributes));
    } catch {}
  }
}

function scopedActivities() {
  const f = CONFIG.fieldMap;
  const role = normalize(attr(state.user, "rol"));
  const region = normalize(attr(state.user, "direccionRegional"));
  const delegation = normalize(attr(state.user, "delegacion"));
  const program = normalize(attr(state.user, "programa"));

  return state.activities.filter(feature => {
    const a = feature.attributes;
    if (role.includes("deleg")) return normalize(a[f.delegacion]) === delegation;
    if (role.includes("regional")) return normalize(a[f.direccionRegional]) === region;
    if (role.includes("coordin")) return normalize(a[f.programa]).includes(program) || program.includes(normalize(a[f.programa]));
    return true;
  });
}

function renderAll() {
  renderKpis();
  renderCharts();
  renderMaps();
  populateFilters();
  renderReviewTable();
  populateFormCatalogs();
  renderNotifications();
  renderNational();
}

function renderKpis() {
  const rows = scopedActivities().map(f => f.attributes);
  const f = CONFIG.fieldMap;
  const total = rows.length;
  const pending = rows.filter(r => isPending(r)).length;
  const approved = rows.filter(r => normalize(r[f.estadoNacional]) === normalize(CONFIG.statuses.approvedNational)).length;
  const participants = rows.reduce((s,r) => s + Number(r[f.participantes] || 0), 0);
  const advance = rows.reduce((s,r) => s + Number(r[f.avance] || 0), 0);

  const data = [
    ["Registros", total],
    ["Pendientes", pending],
    ["Aprobados", approved],
    ["Participantes", participants],
    ["Avance", advance]
  ];
  $("kpiGrid").innerHTML = data.map(([l,v]) => `<div class="kpi"><span>${l}</span><strong>${Number(v).toLocaleString("es-CR")}</strong></div>`).join("");
}

function renderCharts() {
  const rows = scopedActivities().map(f => f.attributes);
  const f = CONFIG.fieldMap;
  const programs = groupCount(rows, f.programa);
  const statuses = groupCount(rows, r => r[f.estadoNacional] || r[f.estadoRegional] || "Sin estado");
  renderBarList("programChart", programs);
  renderBarList("statusChart", statuses);
}

function groupCount(rows, key) {
  const out = {};
  rows.forEach(r => {
    const label = typeof key === "function" ? key(r) : r[key];
    const k = label || "Sin dato";
    out[k] = (out[k] || 0) + 1;
  });
  return Object.entries(out).sort((a,b) => b[1]-a[1]);
}

function renderBarList(id, data) {
  const max = Math.max(1, ...data.map(x => x[1]));
  $(id).innerHTML = data.length ? data.map(([label,value]) =>
    `<div class="chart-row"><span title="${escapeHtml(label)}">${escapeHtml(label)}</span><div class="chart-track"><div class="chart-fill" style="width:${value/max*100}%"></div></div><strong>${value}</strong></div>`
  ).join("") : `<p class="muted">Sin datos.</p>`;
}

function renderMaps() {
  const rows = scopedActivities();
  createMap("dashboardMap", rows, "map");
}

function createMap(containerId, features, stateKey) {
  require(["esri/Map","esri/views/MapView","esri/Graphic","esri/layers/GraphicsLayer"], (Map, MapView, Graphic, GraphicsLayer) => {
    if (state[stateKey]) {
      state[stateKey].destroy();
      state[stateKey] = null;
    }
    const map = new Map({ basemap: "streets-navigation-vector" });
    const layer = new GraphicsLayer();
    map.add(layer);
    features.forEach(feature => {
      const geom = feature.geometry;
      if (!geom) return;
      const a = feature.attributes;
      layer.add(new Graphic({
        geometry: { type:"point", longitude:geom.x, latitude:geom.y, spatialReference:{wkid:4326} },
        symbol: { type:"simple-marker", color:[0,43,127], size:9, outline:{color:[255,255,255],width:1} },
        popupTemplate: {
          title: `{${CONFIG.fieldMap.delegacion}}`,
          content: `<b>Programa:</b> {${CONFIG.fieldMap.programa}}<br><b>Actividad:</b> {${CONFIG.fieldMap.actividad}}<br><b>Estado:</b> {${CONFIG.fieldMap.estadoNacional}}`
        },
        attributes:a
      }));
    });
    const view = new MapView({ container:containerId, map, center:[-84.1,9.95], zoom:7 });
    state[stateKey] = view;
    if (layer.graphics.length) view.goTo(layer.graphics).catch(()=>{});
  });
}

function populateFilters() {
  const rows = scopedActivities().map(f => f.attributes);
  fillSelect("filterRegion", unique(rows, CONFIG.fieldMap.direccionRegional), "Todas las regiones");
  fillSelect("filterDelegation", unique(rows, CONFIG.fieldMap.delegacion), "Todas las delegaciones");
  fillSelect("filterProgram", unique(rows, CONFIG.fieldMap.programa), "Todos los programas");
  fillSelect("filterStatus", [
    CONFIG.statuses.pendingRegional, CONFIG.statuses.approvedRegional, CONFIG.statuses.returnedDelegation,
    CONFIG.statuses.pendingNational, CONFIG.statuses.approvedNational, CONFIG.statuses.rejectedNational
  ], "Todos los estados");
}

function unique(rows, key) {
  return [...new Set(rows.map(r => r[key]).filter(Boolean))].sort();
}

function fillSelect(id, values, firstLabel) {
  const current = $(id).value;
  $(id).innerHTML = `<option value="">${firstLabel}</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (values.includes(current)) $(id).value = current;
}

function reviewCandidates() {
  const role = normalize(attr(state.user, "rol"));
  const f = CONFIG.fieldMap;
  return scopedActivities().filter(feature => {
    const a = feature.attributes;
    if (role.includes("regional")) return [CONFIG.statuses.pendingRegional, CONFIG.statuses.observedRegional].includes(a[f.estadoRegional] || CONFIG.statuses.pendingRegional);
    if (role.includes("coordin")) return normalize(a[f.estadoRegional]) === normalize(CONFIG.statuses.approvedRegional);
    if (role.includes("nacional") || role.includes("admin")) return true;
    return false;
  });
}

function renderReviewTable() {
  const f = CONFIG.fieldMap;
  let rows = reviewCandidates();
  const filters = {
    [f.direccionRegional]: $("filterRegion").value,
    [f.delegacion]: $("filterDelegation").value,
    [f.programa]: $("filterProgram").value
  };
  rows = rows.filter(feature => Object.entries(filters).every(([k,v]) => !v || feature.attributes[k] === v));
  const statusFilter = $("filterStatus").value;
  if (statusFilter) rows = rows.filter(feature => {
    const a = feature.attributes;
    return a[f.estadoNacional] === statusFilter || a[f.estadoRegional] === statusFilter;
  });

  $("reviewTableBody").innerHTML = rows.length ? rows.map(feature => {
    const a = feature.attributes;
    const status = a[f.estadoNacional] || a[f.estadoRegional] || "Pendiente";
    return `<tr>
      <td>${escapeHtml(a[f.objectId])}</td>
      <td>${formatDate(a[f.fechaActividad])}</td>
      <td>${escapeHtml(a[f.direccionRegional])}</td>
      <td>${escapeHtml(a[f.delegacion])}</td>
      <td>${escapeHtml(a[f.programa])}</td>
      <td>${escapeHtml(a[f.actividad])}</td>
      <td><span class="status-pill ${statusClass(status)}">${escapeHtml(status)}</span></td>
      <td><button class="btn btn-secondary btn-review" data-id="${a[f.objectId]}">Revisar</button></td>
    </tr>`;
  }).join("") : `<tr><td colspan="8">No hay registros disponibles.</td></tr>`;

  document.querySelectorAll(".btn-review").forEach(btn => btn.addEventListener("click", () => openReviewModal(btn.dataset.id)));
}

function openReviewModal(objectId) {
  const f = CONFIG.fieldMap;
  const feature = scopedActivities().find(x => String(x.attributes[f.objectId]) === String(objectId));
  if (!feature) return;
  state.selectedActivity = feature;
  const a = feature.attributes;
  const fields = [
    ["ID", a[f.objectId]],["Fecha", formatDate(a[f.fechaActividad])],["Región",a[f.direccionRegional]],["Delegación",a[f.delegacion]],
    ["Programa",a[f.programa]],["Actividad",a[f.actividad]],["Lugar",a[f.lugar]],["Responsable",a[f.responsable]],
    ["Participantes",a[f.participantes]],["Observaciones",a[f.observaciones]]
  ];
  $("reviewDetail").innerHTML = `<div class="detail-grid">${fields.map(([l,v]) => `<div class="detail-item"><small>${l}</small><strong>${escapeHtml(v)}</strong></div>`).join("")}</div>`;
  $("reviewForm").objectid.value = objectId;

  const role = normalize(attr(state.user, "rol"));
  const decisions = role.includes("regional")
    ? [CONFIG.statuses.approvedRegional, CONFIG.statuses.returnedDelegation, CONFIG.statuses.observedRegional]
    : [CONFIG.statuses.approvedNational, CONFIG.statuses.rejectedNational, CONFIG.statuses.observedNational];
  $("reviewForm").decision.innerHTML = decisions.map(x => `<option>${x}</option>`).join("");
  $("reviewModal").classList.remove("hidden");
}

function closeReviewModal() {
  $("reviewModal").classList.add("hidden");
  state.selectedActivity = null;
  $("reviewForm").reset();
}

async function submitReview(event) {
  event.preventDefault();
  const form = new FormData(event.target);
  const decision = form.get("decision");
  const observations = String(form.get("observaciones") || "");
  if ((decision.includes("Devuelta") || decision.includes("observaciones") || decision.includes("Rechazada")) && !observations.trim()) {
    showToast("Debe indicar una observación.", true);
    return;
  }
  const role = normalize(attr(state.user, "rol"));
  const f = CONFIG.fieldMap;
  const updates = {};
  if (role.includes("regional")) {
    updates[f.estadoRegional] = decision;
    updates[f.observacionRegional] = observations;
    if (decision === CONFIG.statuses.approvedRegional) updates[f.estadoNacional] = CONFIG.statuses.pendingNational;
  } else {
    updates[f.estadoNacional] = decision;
    updates[f.observacionNacional] = observations;
  }

  try {
    if (!state.demo) {
      await svc.updateFeature(CONFIG.layers.actividades, form.get("objectid"), updates);
      await writeValidation(form.get("objectid"), decision, observations);
      await writeLog(form.get("objectid"), "REVISIÓN", decision, observations);
      await loadAll();
    } else {
      Object.assign(state.selectedActivity.attributes, updates);
      renderAll();
    }
    closeReviewModal();
    showToast("Revisión guardada.");
  } catch (error) {
    showToast(error.message, true);
  }
}

async function writeValidation(activityId, decision, observations) {
  if (!CONFIG.layers.validaciones) return;
  await svc.addFeature(CONFIG.layers.validaciones, {
    id_actividad: Number(activityId),
    usuario: attr(state.user, "usuario"),
    rol: attr(state.user, "rol"),
    estado: decision,
    observaciones,
    fecha: Date.now()
  });
}

async function writeLog(activityId, action, stateValue, detail) {
  if (!CONFIG.layers.bitacora) return;
  await svc.addFeature(CONFIG.layers.bitacora, {
    id_actividad: Number(activityId),
    usuario: attr(state.user, "usuario"),
    rol: attr(state.user, "rol"),
    accion: action,
    estado_nuevo: stateValue,
    detalle: detail,
    fecha: Date.now()
  });
}

async function submitActivity(event) {
  event.preventDefault();
  const data = Object.fromEntries(new FormData(event.target));
  const f = CONFIG.fieldMap;
  const attributes = {
    [f.fechaActividad]: new Date(`${data.fecha_actividad}T12:00:00`).getTime(),
    [f.horaActividad]: data.hora_actividad,
    [f.direccionRegional]: attr(state.user, "direccionRegional"),
    [f.delegacion]: attr(state.user, "delegacion"),
    [f.programa]: data.programa,
    [f.actividad]: data.actividad,
    [f.provincia]: data.provincia,
    [f.canton]: data.canton,
    [f.distrito]: data.distrito,
    [f.tipoLugar]: data.tipo_lugar,
    [f.lugar]: data.lugar,
    [f.responsable]: data.responsable,
    [f.avance]: Number(data.avance || 0),
    [f.participantes]: Number(data.participantes || 0),
    [f.hombres]: Number(data.hombres || 0),
    [f.mujeres]: Number(data.mujeres || 0),
    [f.edad1018]: Number(data.edad_10_18 || 0),
    [f.edad1930]: Number(data.edad_19_30 || 0),
    [f.edad3145]: Number(data.edad_31_45 || 0),
    [f.edad46Mas]: Number(data.edad_46_mas || 0),
    [f.instituciones]: data.instituciones,
    [f.numeroReferencia]: data.numero_referencia,
    [f.expedienteReferencia]: data.expediente_referencia,
    [f.observaciones]: data.observaciones,
    [f.estadoRegional]: CONFIG.statuses.pendingRegional,
    [f.estadoNacional]: CONFIG.statuses.pendingNational,
    [f.usuarioRegistra]: attr(state.user, "usuario"),
    [f.fechaRegistro]: Date.now()
  };
  const geometry = data.latitud && data.longitud ? { x:Number(data.longitud), y:Number(data.latitud), spatialReference:{wkid:4326} } : null;

  try {
    if (!state.demo) {
      const result = await svc.addFeature(CONFIG.layers.actividades, attributes, geometry);
      await writeLog(result.objectId, "CREACIÓN", CONFIG.statuses.pendingRegional, "Registro creado por delegación");
      await loadAll();
    } else {
      state.activities.push({ attributes:{...attributes,OBJECTID:Date.now()}, geometry });
      renderAll();
    }
    event.target.reset();
    showToast("Actividad registrada y enviada a revisión regional.");
  } catch (error) {
    showToast(error.message, true);
  }
}

async function submitUser(event) {
  event.preventDefault();
  const data = Object.fromEntries(new FormData(event.target));
  const f = CONFIG.fieldMap;
  try {
    if (!state.demo) {
      await svc.addFeature(CONFIG.layers.usuarios, {
        [f.usuario]: data.usuario,
        [f.nombre]: data.nombre,
        [f.rol]: data.rol,
        [f.direccionRegional]: data.direccion_regional,
        [f.delegacion]: data.delegacion,
        [f.programa]: data.programa,
        [f.activo]: 1
      });
    }
    event.target.reset();
    showToast("Usuario autorizado creado.");
  } catch (error) {
    showToast(error.message, true);
  }
}

function populateFormCatalogs() {
  const rows = state.catalogs;
  const programs = [...new Set(rows.map(r => r.programa).filter(Boolean))];
  fillSelect("programSelect", programs.length ? programs : ["DARE","GREAT","MPAS","PSCC","VIF","Política Pública"], "Seleccione");
  populateActivityOptions();

  const provinces = [...new Set(rows.map(r => r.provincia).filter(Boolean))];
  fillSelect("provinceSelect", provinces.length ? provinces : ["San José","Alajuela","Cartago","Heredia","Guanacaste","Puntarenas","Limón"], "Seleccione");
  populateCantons();
}

function populateActivityOptions() {
  const program = $("programSelect").value;
  const activities = [...new Set(state.catalogs.filter(r => !program || r.programa === program).map(r => r.actividad || r.actividad_realizada).filter(Boolean))];
  fillSelect("activitySelect", activities.length ? activities : ["Actividad programada"], "Seleccione");
}

function populateCantons() {
  const province = $("provinceSelect").value;
  const values = [...new Set(state.catalogs.filter(r => !province || r.provincia === province).map(r => r.canton).filter(Boolean))];
  fillSelect("cantonSelect", values.length ? values : ["Cantón"], "Seleccione");
  populateDistricts();
}

function populateDistricts() {
  const province = $("provinceSelect").value;
  const canton = $("cantonSelect").value;
  const values = [...new Set(state.catalogs.filter(r => (!province || r.provincia===province) && (!canton || r.canton===canton)).map(r => r.distrito).filter(Boolean))];
  fillSelect("districtSelect", values.length ? values : ["Distrito"], "Seleccione");
}

function getLocation() {
  if (!navigator.geolocation) return showToast("El navegador no permite geolocalización.", true);
  navigator.geolocation.getCurrentPosition(pos => {
    $("latitudeInput").value = pos.coords.latitude;
    $("longitudeInput").value = pos.coords.longitude;
    showToast("Ubicación obtenida.");
  }, err => showToast(err.message, true), { enableHighAccuracy:true });
}

function renderNotifications() {
  $("notificationBadge").textContent = state.notifications.length;
  $("notificationBadge").classList.toggle("hidden", !state.notifications.length);
  $("notificationList").innerHTML = state.notifications.length ? state.notifications.map(n =>
    `<div class="notification-item"><strong>${escapeHtml(n.mensaje || n.message)}</strong><br><small>${escapeHtml(n.fecha ? formatDate(n.fecha) : "")}</small></div>`
  ).join("") : `<p class="muted">No hay notificaciones pendientes.</p>`;
}

function buildDerivedNotifications() {
  const f = CONFIG.fieldMap;
  const role = normalize(attr(state.user, "rol"));
  const rows = scopedActivities().map(x => x.attributes);
  const out = [];
  if (role.includes("regional")) {
    const count = rows.filter(r => normalize(r[f.estadoRegional] || CONFIG.statuses.pendingRegional) === normalize(CONFIG.statuses.pendingRegional)).length;
    if (count) out.push({ mensaje:`${count} actividad(es) pendiente(s) de revisión regional.`, fecha:Date.now() });
  } else if (role.includes("coordin")) {
    const count = rows.filter(r => normalize(r[f.estadoRegional]) === normalize(CONFIG.statuses.approvedRegional) && normalize(r[f.estadoNacional]) === normalize(CONFIG.statuses.pendingNational)).length;
    if (count) out.push({ mensaje:`${count} actividad(es) pendiente(s) de validación nacional.`, fecha:Date.now() });
  }
  return out;
}

function renderNational() {
  const f = CONFIG.fieldMap;
  const approved = state.activities.filter(x => normalize(x.attributes[f.estadoNacional]) === normalize(CONFIG.statuses.approvedNational));
  const rows = state.summary.length ? state.summary.map(x => x.attributes) : approved.map(x => x.attributes);
  const totals = {
    records: approved.length,
    meta: rows.reduce((s,r)=>s+Number(r[f.meta]||0),0),
    advance: rows.reduce((s,r)=>s+Number(r[f.avance]||0),0),
    participants: approved.reduce((s,x)=>s+Number(x.attributes[f.participantes]||0),0)
  };
  totals.pending = Math.max(totals.meta - totals.advance, 0);
  $("nationalKpis").innerHTML = [
    ["Validadas",totals.records],["Meta",totals.meta],["Avance",totals.advance],["Pendiente",totals.pending],["Participantes",totals.participants]
  ].map(([l,v])=>`<div class="kpi"><span>${l}</span><strong>${Number(v).toLocaleString("es-CR")}</strong></div>`).join("");

  $("nationalTableBody").innerHTML = rows.map(r => {
    const meta = Number(r[f.meta]||0), advance = Number(r[f.avance]||0), pending = Math.max(meta-advance,0), pct = meta ? advance/meta*100 : 0;
    return `<tr><td>${escapeHtml(r[f.direccionRegional])}</td><td>${escapeHtml(r[f.delegacion])}</td><td>${escapeHtml(r[f.programa])}</td><td>${escapeHtml(r[f.actividad])}</td><td>${meta}</td><td>${advance}</td><td>${pending}</td><td>${pct.toFixed(1)}%</td></tr>`;
  }).join("");

  createMap("nationalMap", approved, "nationalMap");
}

function isPending(r) {
  const f = CONFIG.fieldMap;
  const values = [r[f.estadoRegional],r[f.estadoNacional]].map(normalize);
  return values.some(v => v.includes("pendiente"));
}

function statusClass(status) {
  const s = normalize(status);
  if (s.includes("aprob") || s.includes("verificada")) return "status-approved";
  if (s.includes("devuelta") || s.includes("rechazada")) return "status-returned";
  return "status-pending";
}

function formatDate(value) {
  if (!value) return "";
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? String(value) : d.toLocaleDateString("es-CR");
}

function showToast(message, error=false) {
  const el = $("toast");
  el.textContent = message;
  el.style.background = error ? "#B42318" : "#111827";
  el.classList.remove("hidden");
  setTimeout(() => el.classList.add("hidden"), 3200);
}

function seedDemo() {
  const f = CONFIG.fieldMap;
  state.catalogs = [
    {programa:"DARE",actividad:"Charla preventiva",provincia:"San José",canton:"San José",distrito:"Carmen"},
    {programa:"GREAT",actividad:"Sesión educativa",provincia:"San José",canton:"San José",distrito:"Merced"},
    {programa:"PSCC",actividad:"Reunión comunitaria",provincia:"Alajuela",canton:"Alajuela",distrito:"Alajuela"}
  ];
  state.activities = [
    {attributes:{OBJECTID:1,[f.direccionRegional]:"Dirección Regional 1 - San José Central",[f.delegacion]:"D01 Carmen",[f.programa]:"DARE",[f.actividad]:"Charla preventiva",[f.fechaActividad]:Date.now()-86400000,[f.participantes]:42,[f.avance]:1,[f.estadoRegional]:CONFIG.statuses.pendingRegional,[f.estadoNacional]:CONFIG.statuses.pendingNational,[f.lugar]:"Escuela Central"},geometry:{x:-84.078,y:9.936}},
    {attributes:{OBJECTID:2,[f.direccionRegional]:"Dirección Regional 1 - San José Central",[f.delegacion]:"D02 Merced",[f.programa]:"GREAT",[f.actividad]:"Sesión educativa",[f.fechaActividad]:Date.now()-172800000,[f.participantes]:31,[f.avance]:1,[f.estadoRegional]:CONFIG.statuses.approvedRegional,[f.estadoNacional]:CONFIG.statuses.approvedNational,[f.lugar]:"Colegio Merced"},geometry:{x:-84.084,y:9.943}}
  ];
  state.summary = [
    {attributes:{[f.direccionRegional]:"Dirección Regional 1 - San José Central",[f.delegacion]:"D01 Carmen",[f.programa]:"DARE",[f.actividad]:"Charla preventiva",[f.meta]:106,[f.avance]:52}},
    {attributes:{[f.direccionRegional]:"Dirección Regional 1 - San José Central",[f.delegacion]:"D02 Merced",[f.programa]:"GREAT",[f.actividad]:"Sesión educativa",[f.meta]:74,[f.avance]:38}}
  ];
}

