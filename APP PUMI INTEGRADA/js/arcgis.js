import { CONFIG } from "./config.js";

const encode = value => encodeURIComponent(value ?? "");

export class ArcGISService {
  constructor() {
    this.token = sessionStorage.getItem("pumi_token") || "";
  }

  setToken(token) {
    this.token = token || "";
    if (token) sessionStorage.setItem("pumi_token", token);
    else sessionStorage.removeItem("pumi_token");
  }

  async request(url, params = {}, method = "GET") {
    const payload = { f: "json", ...params };
    if (this.token) payload.token = this.token;

    let response;
    if (method === "POST") {
      response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams(payload)
      });
    } else {
      const qs = new URLSearchParams(payload);
      response = await fetch(`${url}?${qs.toString()}`);
    }

    const data = await response.json();
    if (data.error) throw new Error(data.error.message || "Error en ArcGIS");
    return data;
  }

  async query(layerUrl, options = {}) {
    const data = await this.request(`${layerUrl}/query`, {
      where: options.where || "1=1",
      outFields: options.outFields || "*",
      returnGeometry: options.returnGeometry ?? true,
      orderByFields: options.orderByFields || "",
      resultRecordCount: options.resultRecordCount || 2000,
      outSR: 4326
    });
    return data.features || [];
  }

  async addFeature(layerUrl, attributes, geometry = null) {
    const feature = { attributes };
    if (geometry) feature.geometry = geometry;
    const data = await this.request(`${layerUrl}/applyEdits`, {
      adds: JSON.stringify([feature])
    }, "POST");
    const result = data.addResults?.[0];
    if (!result?.success) throw new Error(result?.error?.description || "No se pudo guardar");
    return result;
  }

  async updateFeature(layerUrl, objectId, attributes) {
    const oidField = CONFIG.fieldMap.objectId;
    const data = await this.request(`${layerUrl}/applyEdits`, {
      updates: JSON.stringify([{ attributes: { [oidField]: objectId, ...attributes } }])
    }, "POST");
    const result = data.updateResults?.[0];
    if (!result?.success) throw new Error(result?.error?.description || "No se pudo actualizar");
    return result;
  }

  async getPortalUser() {
    if (!this.token) return null;
    return this.request(`${CONFIG.portalUrl}/sharing/rest/community/self`, {});
  }

  async getAppUser(username) {
    const f = CONFIG.fieldMap;
    const where = `${f.usuario}='${String(username).replaceAll("'", "''")}' AND (${f.activo}=1 OR ${f.activo}='1' OR ${f.activo}='true')`;
    const rows = await this.query(CONFIG.layers.usuarios, { where, returnGeometry: false });
    return rows[0]?.attributes || null;
  }
}

export function startOAuth() {
  const state = crypto.randomUUID();
  sessionStorage.setItem("pumi_oauth_state", state);
  const authorize = `${CONFIG.portalUrl}/sharing/rest/oauth2/authorize`;
  const params = new URLSearchParams({
    client_id: CONFIG.oauthClientId,
    response_type: "token",
    expiration: "120",
    redirect_uri: CONFIG.oauthRedirectUri,
    state
  });
  window.location.assign(`${authorize}?${params}`);
}

export function consumeOAuthHash() {
  if (!location.hash.includes("access_token")) return "";
  const params = new URLSearchParams(location.hash.substring(1));
  const token = params.get("access_token") || "";
  history.replaceState({}, document.title, location.pathname + location.search);
  return token;
}

