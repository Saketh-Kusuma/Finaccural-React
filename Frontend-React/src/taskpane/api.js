export const API_BASE = "http://localhost:8000";

export async function apiFetch(path, options = {}) {
  const token = localStorage.getItem("fa_jwt_token");
  const headers = { ...(options.headers || {}) };
  if (token) headers.Authorization = `Bearer ${token}`;
  if (options.body && !headers["Content-Type"]) headers["Content-Type"] = "application/json";
  const response = await fetch(path.startsWith("http") ? path : `${API_BASE}${path}`, { ...options, headers });
  if (!response.ok) throw new Error((await response.json().catch(() => ({}))).message || "Request failed.");
  return response;
}

export function openAuth(provider, onProfile, loginHint) {
  let url = `${API_BASE}/api/${provider === "google" ? "auth/google" : "microsoft"}/connect`;
  if (loginHint) url += `?login_hint=${encodeURIComponent(loginHint)}`;
  const popup = window.open(url, `fa_${provider}_auth`, "width=520,height=640,resizable=yes,scrollbars=yes");
  if (!popup) throw new Error("The sign-in window was blocked. Please allow popups and try again.");
  const receive = (event) => {
    const data = event.data || {};
    const types = provider === "google" ? ["google_authed", "google_profile"] : ["microsoft_authed", "ms_authed", "microsoft_profile", "ms_profile"];
    if (!types.includes(data.type)) return;
    window.removeEventListener("message", receive);
    onProfile({ ...data, provider });
  };
  window.addEventListener("message", receive);
}

export function openErp(provider, user) {
  const path = provider === "quickbooks" ? "/api/quickbooks/connect/" : "/api/xero/connect";
  const params = new URLSearchParams({ tier: user.plan || "", mail: user.email || "", token: localStorage.getItem("fa_jwt_token") || "" });
  return window.open(`${API_BASE}${path}?${params}`, `fa_${provider}_auth`, "width=800,height=600,resizable=yes,scrollbars=yes");
}

export function openTrialSelectDialog(onAction) {
  const dialogUrl = window.location.origin + "/trialselect.html";

  // If running inside Office Add-in host with dialog API available
  if (
    typeof Office !== "undefined" &&
    Office.context &&
    Office.context.ui &&
    Office.context.ui.displayDialogAsync
  ) {
    let dialog = null;
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 45, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.warn("displayDialogAsync failed, falling back to popup:", asyncResult.error.message);
          openTrialPopupFallback(dialogUrl, onAction);
        } else {
          dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            try {
              const message = typeof arg.message === "string" ? JSON.parse(arg.message) : arg.message;
              if (dialog) dialog.close();
              if (message?.type === "START_TRIAL") onAction("START_TRIAL");
              if (message?.type === "VIEW_PLANS") onAction("VIEW_PLANS");
            } catch (err) {
              console.error("Error parsing dialog message:", err);
            }
          });
        }
      }
    );
    return;
  }

  // Fallback if not inside Office or displayDialogAsync not available
  openTrialPopupFallback(dialogUrl, onAction);
}

function openTrialPopupFallback(url, onAction) {
  const popup = window.open(
    url,
    "fa_trialselect",
    "width=640,height=560,resizable=yes,scrollbars=yes"
  );
  const receive = (event) => {
    let data = event.data;
    if (typeof data === "string") {
      try { data = JSON.parse(data); } catch (_) {}
    }
    if (data?.type === "START_TRIAL" || data?.type === "VIEW_PLANS") {
      window.removeEventListener("message", receive);
      if (popup && !popup.closed) popup.close();
      onAction(data.type);
    }
  };
  window.addEventListener("message", receive);
}

export async function fetchMasterDataStream(provider, companyId, tier, onProgress) {
  const params = new URLSearchParams({
    companyId: companyId || "",
    platform: provider || "",
    tier: tier || "",
    stream: "true"
  });

  const response = await apiFetch(`/api/pull-master-data?${params.toString()}`, { method: "GET" });
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let finalData = null;

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    const lines = buffer.split("\n\n");
    buffer = lines.pop();

    for (const line of lines) {
      if (line.startsWith("data: ")) {
        try {
          const payload = JSON.parse(line.replace("data: ", "").trim());
          if (payload.type === "progress" && typeof onProgress === "function") {
            onProgress(payload);
          } else if (payload.type === "start" && typeof onProgress === "function") {
            onProgress(payload);
          } else if (payload.type === "complete") {
            finalData = payload.data;
          } else if (payload.type === "error") {
            throw new Error(payload.error || "Data sync error");
          }
        } catch (pErr) {
          if (pErr.message && !pErr.message.includes("Unexpected token")) throw pErr;
        }
      }
    }
  }

  return finalData;
}

export async function fetchIncrementalDataStream(provider, companyId, tier, onProgress) {
  const params = new URLSearchParams({
    companyId: companyId || "",
    platform: provider || "",
    tier: tier || "",
    stream: "true",
    mode: "incremental"
  });

  const response = await apiFetch(`/api/pull-master-data?${params.toString()}`, { method: "GET" });
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let finalData = null;

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    const lines = buffer.split("\n\n");
    buffer = lines.pop();

    for (const line of lines) {
      if (line.startsWith("data: ")) {
        try {
          const payload = JSON.parse(line.replace("data: ", "").trim());
          if (payload.type === "progress" && typeof onProgress === "function") {
            onProgress(payload);
          } else if (payload.type === "start" && typeof onProgress === "function") {
            onProgress(payload);
          } else if (payload.type === "complete") {
            finalData = payload.data;
          } else if (payload.type === "error") {
            throw new Error(payload.error || "Incremental sync error");
          }
        } catch (pErr) {
          if (pErr.message && !pErr.message.includes("Unexpected token")) throw pErr;
        }
      }
    }
  }

  return finalData;
}

