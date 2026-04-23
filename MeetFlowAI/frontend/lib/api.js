/**
 * Base da API.
 * - Com proxy do Next (recomendado no dev): prefixo /__meetflow (mesma origem, sem CORS).
 * - Com NEXT_PUBLIC_API_URL: URL absoluta (ex.: produção ou acesso direto).
 */
function getApiBase() {
  const fromEnv = process.env.NEXT_PUBLIC_API_URL;
  if (fromEnv) return fromEnv.replace(/\/$/, "");

  if (typeof window !== "undefined") {
    return "/__meetflow";
  }
  // SSR / build: requisições do servidor vão direto ao backend (testes, etc.)
  return process.env.INTERNAL_API_URL || "http://127.0.0.1:8000";
}

function buildUrl(path) {
  const base = getApiBase();
  if (path.startsWith("http")) return path;
  return `${base}${path}`;
}

function formatApiError(data) {
  if (!data) return "Falha no pedido";
  if (typeof data.detail === "string") return data.detail;
  if (typeof data.detail === "object" && data.detail?.message) return data.detail.message;
  if (Array.isArray(data.detail)) {
    return data.detail
      .map((d) => (typeof d === "string" ? d : d.msg || JSON.stringify(d)))
      .join(" ");
  }
  if (data.error) return typeof data.error === "string" ? data.error : JSON.stringify(data.error);
  return "Falha no pedido";
}

async function request(path, options = {}, token = "") {
  const url = buildUrl(path);
  let response;
  try {
    response = await fetch(url, {
      ...options,
      headers: {
        ...(options.headers || {}),
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
    });
  } catch (err) {
    if (err?.name === "AbortError") {
      throw err;
    }
    const tip =
      "Não foi possível conectar ao backend. Confirme se o FastAPI está em execução (ex.: uvicorn em 0.0.0.0:8000) e, no navegador, use o proxy /__meetflow do Next em desenvolvimento.";
    const netErr = new Error(err?.message ? `${err.message}. ${tip}` : tip);
    netErr.status = 0;
    throw netErr;
  }
  if (response.status === 204) {
    return null;
  }
  const contentType = response.headers.get("content-type") || "";
  const data = contentType.includes("application/json") ? await response.json() : await response.text();
  if (!response.ok) {
    const e = new Error(formatApiError(typeof data === "object" ? data : { detail: String(data) }));
    e.status = response.status;
    const retryHeader = response.headers.get("retry-after");
    const retryFromBody =
      typeof data === "object" && typeof data?.detail === "object" ? Number(data.detail?.retry_after_seconds || 0) : 0;
    e.retryAfterSeconds = Number(retryHeader || 0) || retryFromBody || 0;
    throw e;
  }
  return data;
}

export const api = {
  login: (email, password) =>
    request("/api/auth/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    }),
  register: (payload) =>
    request("/api/auth/register", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    }),
  me: (token) => request("/api/me", {}, token),
  updateMe: (payload, token) =>
    request(
      "/api/me",
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      },
      token,
    ),
  changePassword: (payload, token) =>
    request(
      "/api/me/change-password",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      },
      token,
    ),
  dashboard: (token) => request("/api/dashboard", {}, token),
  meetings: (token) => request("/api/meetings", {}, token),
  meetingDetail: (id, token) => request(`/api/meetings/${id}`, {}, token),
  processText: (payload, token, signal) =>
    request(
      "/api/meetings/process-text",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
        signal,
      },
      token,
    ),
  processUpload: (formData, token, signal) =>
    request("/api/meetings/process-upload", { method: "POST", body: formData, signal }, token),
  deleteMeeting: (id, token) => request(`/api/meetings/${id}`, { method: "DELETE" }, token),
  chat: (id, question, token, signal) =>
    request(
      `/api/meetings/${id}/chat`,
      { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ question }), signal },
      token,
    ),
  exportMeeting: async (id, format, token) => {
    const url = buildUrl(`/api/meetings/${id}/export/${format}`);
    let response;
    try {
      response = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
      });
    } catch (err) {
      throw new Error("Não foi possível ligar ao backend para exportar.");
    }
    if (!response.ok) throw new Error("Falha ao exportar");
    return response.blob();
  },
};
