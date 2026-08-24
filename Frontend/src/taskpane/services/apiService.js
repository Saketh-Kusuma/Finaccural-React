/**
 * Backend API layer — every `fetch()` in the add-in goes through here.
 *
 * `apiFetch()` is the single choke point that attaches the Authorization
 * header, transparently refreshes an expired access token (queueing any
 * concurrent 401s behind one refresh call) and normalises failures into
 * ApiError instances with a stable `.code`.
 */
import { AppState } from "../state/appState.js";
import { AuthService } from "./authService.js";
import { DashboardService } from "./dashboardService.js";
import {
    ERROR_CODES,
    ApiError,
    parseApiError,
    networkError,
    showBanner,
    hideBanner,
    showToast
} from "../../shared/apiErrorHandler.js";
import { getFriendlyMessage } from "../../shared/errorMessages.js";

const ApiService = {
    // ── Single source-of-truth for the backend base URL ────────────
    BASE: "http://localhost:8000",

    /**
     * Centralized error reaction — every call through apiFetch funnels
     * here. Branches on the standardized backend `code` field (never on
     * message text) and reacts globally:
     *   - ERR_CONNECTION_REFUSED  -> red offline banner + Retry
     *   - ERR_SESSION_EXPIRED     -> toast + clear tokens + redirect to
     *                                 Login + block further requests
     *   - ERR_ERP_SESSION_EXPIRED -> orange banner + Reconnect
     * Any other code is left to the caller's own try/catch — this only
     * reacts to the three centrally-handled scenarios.
     * @param {ApiError} apiErr
     * @param {{ retry?: () => void }} [opts]
     */
    handleGlobalApiError(apiErr, opts = {}) {
        switch (apiErr.code) {
            case ERROR_CODES.CONNECTION_REFUSED: {
                showBanner({
                    type: "offline",
                    message: apiErr.message,
                    actionLabel: "Retry",
                    onAction: () => {
                        hideBanner();
                        if (opts.retry) opts.retry();
                    }
                });
                break;
            }
            case ERROR_CODES.SESSION_EXPIRED: {
                // Avoid re-triggering the redirect/toast for every
                // in-flight request that fails after the first 401.
                if (AppState.sessionExpired) break;
                AppState.sessionExpired = true;
                showToast(getFriendlyMessage(ERROR_CODES.SESSION_EXPIRED));
                AuthService.logout();
                break;
            }
            case ERROR_CODES.ERP_SESSION_EXPIRED: {
                const provider = AppState.currentProvider === "xero" ? "Xero" : "QuickBooks";
                showBanner({
                    type: "erp",
                    message: apiErr.message,
                    actionLabel: "Reconnect",
                    onAction: () => {
                        hideBanner();
                        DashboardService.launchERPOAuth(AppState.currentProvider === "xero" ? "xero" : "quickbooks");
                    }
                });
                void provider; // reserved for future provider-specific copy
                break;
            }
            case ERROR_CODES.QB_SUBSCRIPTION_EXPIRED: {
                showBanner({
                    type: "erp",
                    message: apiErr.message || "Your QuickBooks subscription has expired or been suspended. Please log into QuickBooks to update your billing.",
                    actionLabel: "Dismiss",
                    onAction: () => {
                        hideBanner();
                    }
                });
                break;
            }
            default:
                // Not one of the three centrally-handled scenarios —
                // caller's own catch block is responsible for the UI.
                break;
        }
    },

    /**
     * Authenticated fetch helper — the single choke point every API call
     * in this app goes through. Automatically attaches the JWT Bearer
     * token, and centrally reacts to the app's three standardized error
     * scenarios (offline/unreachable backend, expired session, expired
     * ERP connection) via handleGlobalApiError above, in addition to
     * returning the raw Response so existing callers keep working
     * unchanged (their own res.ok / res.json() handling still applies).
     *
     * Usage (same API as window.fetch):
     *   const res = await ApiService.apiFetch('/api/connections', { method: 'GET' });
     *
     * @param {string}  path     - Relative (/api/...) or absolute URL
     * @param {object}  options  - Standard fetch options (method, body, headers, …)
     * @returns {Promise<Response>}
     */
    async apiFetch(path, options = {}) {
        // Once a session-expired redirect has fired, refuse further
        // requests until a fresh token is obtained via login — prevents
        // a storm of repeated 401s while the user is on the Login view.
        if (AppState.sessionExpired) {
            throw new ApiError(ERROR_CODES.SESSION_EXPIRED, getFriendlyMessage(ERROR_CODES.SESSION_EXPIRED), "Blocked: session already marked expired.");
        }

        const url = path.startsWith("http") ? path : `${this.BASE}${path}`;
        const headers = { ...(options.headers || {}) };
        if (AppState.jwtToken) {
            headers["Authorization"] = `Bearer ${AppState.jwtToken}`;
        }

        const retry = () => this.apiFetch(path, options);

        let res;
        try {
            res = await fetch(url, { ...options, headers });
        } catch (networkErr) {
            const apiErr = networkError(networkErr);
            this.handleGlobalApiError(apiErr, { retry });
            throw apiErr;
        }

        if (!res.ok) {
            // ── Auto-refresh on 401 ─────────────────────────────────
            // Only attempt once (no infinite loop) and only if we have
            // a stored refresh token. A 401 can mean:
            //   a) The 15-min access JWT expired  → refresh and retry
            //   b) The refresh token itself is bad → fall through to
            //      the existing ERR_SESSION_EXPIRED path.
            const parsed = await parseApiError(res.clone());
            if (parsed.status === 401 && AppState.refreshToken && !options._retried) {
                try {
                    // Serialize concurrent 401s — only one refresh call
                    if (AppState._refreshing) {
                        // Queue this retry behind the in-flight refresh
                        await new Promise((resolve, reject) =>
                            AppState._refreshQueue.push({ resolve, reject })
                        );
                        // Refresh completed — retry with the new token
                        return this.apiFetch(path, { ...options, _retried: true });
                    }
                    AppState._refreshing = true;
                    const refreshRes = await fetch(`${this.BASE}/api/auth/refresh`, {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ refreshToken: AppState.refreshToken })
                    });
                    if (refreshRes.ok) {
                        const data = await refreshRes.json();
                        AppState.jwtToken = data.token;
                        AppState.refreshToken = data.refreshToken;
                        AppState.sessionExpired = false;
                        localStorage.setItem("fa_jwt_token", data.token);
                        localStorage.setItem("fa_refresh_token", data.refreshToken);
                        // Unblock queued retries
                        AppState._refreshQueue.forEach(p => p.resolve());
                        AppState._refreshQueue = [];
                        AppState._refreshing = false;
                        return this.apiFetch(path, { ...options, _retried: true });
                    } else {
                        // Refresh failed — fall through to session-expired
                        AppState._refreshQueue.forEach(p => p.reject());
                        AppState._refreshQueue = [];
                        AppState._refreshing = false;
                    }
                } catch (_) {
                    AppState._refreshing = false;
                    AppState._refreshQueue = [];
                }
            }
            // ── Normal error path ────────────────────────────────────
            this.handleGlobalApiError(parsed, { retry });
            return res;
        }

        // A successful call clears any stale offline banner.
        hideBanner();
        return res;
    },

    /**
     * Checks subscription status for the given email from backend.
     * Also stores the JWT token if the server returns one.
     *
     * If the backend is unreachable (no internet / server down), this
     * surfaces the same offline banner + Retry action as every other
     * call in the app instead of failing silently — the caller still
     * gets a graceful fallback value so existing view-routing logic
     * keeps working, but the user now sees *why*.
     */
    async checkSubscription(email) {
        try {
            const res = await this.apiFetch(`/api/auth/me`);
            const result = await res.json();
            hideBanner();
            const user = result.user || {};
            return {
                hasSubscription: !!user.plan,
                plan: user.plan,
                subscriptionId: user.subscriptionId,
                user: user,
                success: true
            };
        } catch (err) {
            // Return fallback state if network error or session expired
            return { hasSubscription: AppState.hasSubscription };
        }
    },

    /**
     * Starts the free trial for the signed-in user (explicit choice
     * from the Free Trial vs Subscription Plan screen). The backend
     * sets plan + a fresh trial_ends_at clock starting now, so the
     * 2-minute countdown begins at the moment the user actually opts
     * in — not at account-creation time.
     * @returns {Promise<{success: boolean, user?: object}>}
     */
    async startTrial() {
        const res = await this.apiFetch("/api/auth/start-trial", { method: "POST" });
        if (!res.ok) throw new Error("Failed to start free trial.");
        return await res.json();
    },

    /** Checks ERP token validity from backend (JWT required). */
    async checkTokens(provider) {
        const path = provider === "quickbooks"
            ? "/api/quickbooks/tokens/"
            : "/api/xero/tokens";
        const res = await this.apiFetch(path);
        if (!res.ok) throw new Error(`Failed to check ${provider} tokens.`);
        return await res.json();
    },

    /** Disconnects the ERP provider from backend (JWT required). */
    async disconnectERP(provider) {
        const path = provider === "quickbooks"
            ? "/api/quickbooks/disconnect"
            : "/api/xero/disconnect";
        const res = await this.apiFetch(path, { method: "POST" });
        if (!res.ok) throw new Error(`Failed to disconnect ${provider}.`);
        return await res.json();
    },

    /**
     * Pulls master metadata from ERP APIs via the unified backend pull endpoint.
     * @param {string} provider  - The active ERP provider ("quickbooks" | "xero")
     * @param {string} companyId - Selected company identifier
     * @param {object|null} [cursor] - Per-entity pagination cursor
     *   returned as `cursor` on a previous call, or omit/null to start
     *   a fresh pull cycle. The backend fetches exactly ONE page (up
     *   to 100 records) of exactly ONE entity for THIS call only — the
     *   entities are drained one at a time, in the fixed order
     *   Accounts -> Classes -> Locations -> Customers -> Vendors — so
     *   four of the five record arrays come back empty on any given
     *   call, and the whole dataset is never fetched at once.
     * @returns {Promise<object>} Map of company, customers, vendors, accounts, classes, locations, plus `cursor` (pass to the next call to continue) and `isDone` (true once every entity is exhausted).
     */
    async fetchMasterDataStream(provider, companyId, onProgress) {
        const params = new URLSearchParams({
            companyId: companyId || "",
            platform: provider || "",
            tier: AppState.currentTier || "",
            stream: "true"
        });

        const url = `${this.BASE}/api/pull-master-data?${params.toString()}`;
        const headers = {};
        if (AppState.jwtToken) {
            headers["Authorization"] = `Bearer ${AppState.jwtToken}`;
        }

        const response = await fetch(url, { method: "GET", headers });
        if (!response.ok) {
            throw await parseApiError(response);
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";
        let finalData = null;

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });

            const lines = buffer.split("\n\n");
            buffer = lines.pop(); // Keep last partial chunk

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
    },

    async fetchMasterData(provider, companyId, cursor) {
        // const apiErr = { code: ERROR_CODES.QB_SUBSCRIPTION_EXPIRED, message: "Your QuickBooks subscription has expired" };
        // ApiService.handleGlobalApiError(apiErr);
        // throw apiErr;
        const params = new URLSearchParams({
            companyId: companyId || "",
            platform: provider || "",
            tier: AppState.currentTier || ""
        });
        if (cursor) {
            params.set("cursor", JSON.stringify(cursor));
        }
        const res = await this.apiFetch(`/api/pull-master-data?${params.toString()}`, {
            method: "GET"
        });
        if (!res.ok) {
            // apiFetch already triggered the global reaction (offline
            // banner / ERP-expired banner) as a side effect above.
            // Re-parse here too (parseApiError clones internally, so
            // this doesn't double-consume the body) so this call's own
            // catch block can branch on `.code` instead of guessing
            // from message text.
            throw await parseApiError(res);
        }
        const data = await res.json();

        if (data && data.tokenRefreshed) {
            DashboardService.addLog("Token is refreshed");
        }

        // This request/response happens entirely in the browser (fetch
        // API running in the taskpane WebView), so the count is logged
        // to the browser console, not a terminal.
        const counts = {
            company: Array.isArray(data.company) ? data.company.length : (data.company ? 1 : 0),
            accounts: Array.isArray(data.accounts) ? data.accounts.length : 0,
            classes: Array.isArray(data.classes) ? data.classes.length : 0,
            locations: Array.isArray(data.locations) ? data.locations.length : 0,
            customers: Array.isArray(data.customers) ? data.customers.length : 0,
            vendors: Array.isArray(data.vendors) ? data.vendors.length : 0
        };
        const total = counts.company + counts.accounts + counts.classes + counts.locations + counts.customers + counts.vendors;
        console.log(
            `[FinAccrual] Master data records for company ${companyId} (${provider}): ` +
            `accounts=${counts.accounts}, classes=${counts.classes}, locations=${counts.locations}, ` +
            `customers=${counts.customers}, vendors=${counts.vendors}, company=${counts.company}, total=${total}`
        );

        return data;
    }
};

export { ApiService };
