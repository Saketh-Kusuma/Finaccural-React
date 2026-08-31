/**
 * Google / Microsoft OAuth popups, session bootstrap and logout.
 *
 * Both popup flows post their result back with window.postMessage; the
 * handlers below own the JWT/refresh-token hand-off into AppState and the
 * new-user vs returning-user branch that follows it.
 */
import { AppState } from "../state/appState.js";
import { ViewRouter } from "../router/viewRouter.js";
import { ApiService } from "./apiService.js";
import { ExcelService } from "./excelService.js";
import { NotificationService } from "./notificationService.js";
import { DashboardService } from "./dashboardService.js";
import { AppController } from "../controllers/appController.js";
import { hideBanner } from "../../shared/apiErrorHandler.js";

const AuthService = {
    /**
     * Called when a new-user completes payment inside the popup.
     * Receives full profile + subscription info from google_authed postMessage.
     * @param {string} email
     * @param {string} name
     * @param {string} provider
     * @param {string} subscriptionId
     * @param {string} plan
     */

    _saveAccountToHistory(account) {
        if (!account || !account.email) return;
        let accounts = [];
        try {
            accounts = JSON.parse(localStorage.getItem("fa_accounts_history") || "[]");
        } catch (e) {
            accounts = [];
        }
        if (!Array.isArray(accounts)) accounts = [];

        accounts = accounts.filter(a => a && a.email && a.email.toLowerCase() !== account.email.toLowerCase());
        accounts.unshift({
            name: account.name || account.email,
            email: account.email,
            provider: account.provider || "google"
        });
        if (accounts.length > 5) accounts = accounts.slice(0, 5);
        localStorage.setItem("fa_accounts_history", JSON.stringify(accounts));
    },

    handleNewUserAuthed(email, name, provider, subscriptionId, plan, token, refreshToken) {
        // Drop any notification history cached in memory for whoever
        // was previously signed in on this taskpane session (e.g. an
        // account switch without a full logout) — otherwise the badge
        // could flash the previous user's unread count for the instant
        // before NotificationService.init()'s backend refetch (below,
        // via DashboardService.render()) lands.
        NotificationService._cache = [];
        NotificationService._lastNotif = null;

        AppState.userEmail = email;
        AppState.userName = name;
        AppState.userProvider = provider;
        AppState.hasSubscription = true;
        AppState.subscriptionId = subscriptionId;
        AppState.subscriptionPlan = plan;

        // Persist the JWT + refresh token so all subsequent API calls are authenticated
        if (token) {
            AppState.jwtToken = token;
            AppState.sessionExpired = false; // fresh token — lift the request block
            localStorage.setItem("fa_jwt_token", token);
        }
        if (refreshToken) {
            AppState.refreshToken = refreshToken;
            localStorage.setItem("fa_refresh_token", refreshToken);
        }

        localStorage.setItem("fa_user_email", email);
        localStorage.setItem("fa_user_name", name);
        localStorage.setItem("fa_user_provider", provider);
        this._saveAccountToHistory({ email, name, provider });
        this._persistSubscription();

        DashboardService.render();
        ViewRouter.show("Dashboard");
        DashboardService.showStatus("Login successful.", "success");

        const modal = document.getElementById("trialExpiredModal");
        if (modal) modal.style.display = "none";

        AppController.startTrialExpirationWatcher();
        this.startTokenRefreshTimer();
    },

    /**
     * Called when returning user signs in (popup closes immediately with google_profile)
     * and backend confirms their subscription.
     * @param {string} email
     * @param {string} name
     * @param {string} provider
     */
    async handleReturningUser(email, name, provider, token, refreshToken) {
        // Same reasoning as handleNewUserAuthed above — clear the
        // previous account's cached notifications before this account's
        // data replaces it.
        NotificationService._cache = [];
        NotificationService._lastNotif = null;

        AppState.userEmail = email;
        AppState.userName = name;
        AppState.userProvider = provider;
        localStorage.setItem("fa_user_email", email);
        localStorage.setItem("fa_user_name", name);
        localStorage.setItem("fa_user_provider", provider);
        this._saveAccountToHistory({ email, name, provider });

        // Persist token from popup if provided (avoids a round-trip)
        if (token) {
            AppState.jwtToken = token;
            AppState.sessionExpired = false; // fresh token — lift the request block
            localStorage.setItem("fa_jwt_token", token);
        }
        if (refreshToken) {
            AppState.refreshToken = refreshToken;
            localStorage.setItem("fa_refresh_token", refreshToken);
        }

        ViewRouter.show("Loading");
        try {
            const result = await ApiService.checkSubscription(email);
            const userPlan = result.plan || (result.user && result.user.plan);
            if (result.hasSubscription || userPlan) {
                AppState.hasSubscription = true;
                AppState.subscriptionId = result.subscriptionId || AppState.subscriptionId;
                AppState.subscriptionPlan = userPlan || AppState.subscriptionPlan;
                this._persistSubscription();
                DashboardService.render();
                ViewRouter.show("Dashboard");
                DashboardService.showStatus("Login successful.", "success");

                const modal = document.getElementById("trialExpiredModal");
                if (modal) modal.style.display = "none";

                AppController.startTrialExpirationWatcher();
                this.startTokenRefreshTimer();
            } else {
                // Plan is null or missing — show trial vs subscribe popup
                AppController.openTrialSelectDialog();
            }
        } catch {
            AppController.openTrialSelectDialog();
            DashboardService.showStatus("Login failed. Please try again.", "error");
        }
    },

    /**
     * Opens a Google OAuth popup.
     * The popup now hosts the entire Plans → Payment → Success flow.
     * - For new users:      popup sends  { type: 'google_authed', email, name, subscriptionId, plan }
     * - For returning users (if future backend check skips popup): sends { type: 'google_profile', ... }
     * - On logout click:    popup sends  { type: 'google_cancelled' }
     *
     * @param {string} [loginHint] - Email of a remembered account
     *   (from the account picker's "previous account" row). When set,
     *   Google skips its own account-chooser screen and goes straight
     *   to that account, landing the user on the Dashboard the same
     *   way an already-connected QuickBooks/Xero company just shows
     *   its details instead of re-prompting to connect.
     */
    openGooglePopup(loginHint) {
        const googleAuthUrl = loginHint
            ? `${ApiService.BASE}/api/auth/google/connect?login_hint=${encodeURIComponent(loginHint)}`
            : `${ApiService.BASE}/api/auth/google/connect`;

        const msgHandler = (event) => {
            if (!event.data) return;
            let data = event.data;
            if (typeof data === "string") {
                try { data = JSON.parse(data); } catch (_) { }
            }

            if (!data || !data.type) return;

            if (data.type === "google_authed") {
                // New user completed payment inside popup
                window.removeEventListener("message", msgHandler);
                AuthService.handleNewUserAuthed(
                    data.email || "",
                    data.name || data.email || "",
                    "google",
                    data.subscriptionId || "",
                    data.plan || "Starter",
                    data.token || "",
                    data.refreshToken || ""
                );
            } else if (data.type === "google_profile") {
                // Returning user — popup closed immediately, check backend
                window.removeEventListener("message", msgHandler);
                AuthService.handleReturningUser(
                    data.email || "",
                    data.name || data.email || "",
                    "google",
                    data.token || "",
                    data.refreshToken || ""
                );
            } else if (data.type === "google_cancelled") {
                // User clicked logout in the popup
                window.removeEventListener("message", msgHandler);
                ViewRouter.show("Welcome");
            }
        };
        window.addEventListener("message", msgHandler);

        const popup = window.open(
            googleAuthUrl, "fa_google_auth",
            "width=640,height=840,top=40,left=80,toolbar=no,menubar=no"
        );

        if (!popup || popup.closed) {
            window.removeEventListener("message", msgHandler);
            DashboardService.showError("Popup was blocked. Please allow popups and try again.");
        } else {
            // Bring the popup to the front — without this it can open
            // behind the taskpane/main window in some browsers.
            popup.focus();
        }
    },

    /**
     * Opens a mock Microsoft OAuth flow.
     * In production, replace with real Microsoft MSAL / OAuth URL.
     *
     * @param {string} [loginHint] - Email of a remembered account; see
     *   openGooglePopup() above for why this is passed through.
     */
    openMicrosoftPopup(loginHint) {
        const mockMSUrl = loginHint
            ? `${ApiService.BASE}/api/microsoft/connect?login_hint=${encodeURIComponent(loginHint)}`
            : `${ApiService.BASE}/api/microsoft/connect`;

        const msgHandler = (event) => {
            if (!event.data) return;
            let data = event.data;
            if (typeof data === "string") {
                try { data = JSON.parse(data); } catch (_) { }
            }
            if (!data || !data.type) return;

            if (data.type === "microsoft_authed" || data.type === "ms_authed") {
                window.removeEventListener("message", msgHandler);
                AuthService.handleNewUserAuthed(
                    data.email || "",
                    data.name || data.email || "",
                    "microsoft",
                    data.subscriptionId || "",
                    data.plan || "Starter",
                    data.token || "",
                    data.refreshToken || ""
                );
            } else if (data.type === "ms_profile" || data.type === "microsoft_profile") {
                window.removeEventListener("message", msgHandler);
                AuthService.handleReturningUser(
                    data.email || "",
                    data.name || data.email || "",
                    "microsoft",
                    data.token || "",
                    data.refreshToken || ""
                );
            } else if (data.type === "ms_cancelled" || data.type === "google_cancelled") {
                window.removeEventListener("message", msgHandler);
                ViewRouter.show("Welcome");
            }
        };
        window.addEventListener("message", msgHandler);

        const popup = window.open(
            mockMSUrl, "fa_ms_auth",
            "width=640,height=800,top=60,left=80,toolbar=no,menubar=no"
        );

        if (!popup || popup.closed) {
            window.removeEventListener("message", msgHandler);
            DashboardService.showError("Popup was blocked. Please allow popups and try again.");
        } else {
            // Bring the popup to the front — without this it can open
            // behind the taskpane/main window in some browsers.
            popup.focus();
        }
    },

    _tokenRefreshInterval: null,

    /**
     * Check if the access JWT will expire within the buffer window (default 5 minutes).
     * @param {string} token - JWT Access Token
     * @param {number} bufferMs - Buffer time in ms (default 5 minutes = 300,000 ms)
     * @returns {boolean}
     */
    isTokenExpiringSoon(token, bufferMs = 5 * 60 * 1000) {
        if (!token) return true;
        try {
            const payloadBase64 = token.split(".")[1];
            const decoded = JSON.parse(atob(payloadBase64.replace(/-/g, "+").replace(/_/g, "/")));
            if (decoded && decoded.exp) {
                const expMs = decoded.exp * 1000;
                return (expMs - Date.now()) <= bufferMs;
            }
        } catch (_) { }
        return true;
    },

    async ensureValidToken() {
        if (!AppState.refreshToken) return AppState.jwtToken;
        if (!this.isTokenExpiringSoon(AppState.jwtToken, 5 * 60 * 1000)) {
            return AppState.jwtToken;
        }
        try {
            const refreshRes = await fetch(`${ApiService.BASE}/api/auth/refresh`, {
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
            }
        } catch (err) {
            console.error("Token refresh failed during pre-flight check:", err);
        }
        return AppState.jwtToken;
    },

    startTokenRefreshTimer() {
        if (this._tokenRefreshInterval) {
            clearInterval(this._tokenRefreshInterval);
        }
        // Checks every 30 seconds if the access token is within 5 minutes of expiration.
        // If remaining time <= 5 minutes, it automatically rotates tokens.
        this._tokenRefreshInterval = setInterval(async () => {
            if (AppState.jwtToken && AppState.refreshToken && !AppState.sessionExpired) {
                if (this.isTokenExpiringSoon(AppState.jwtToken, 5 * 60 * 1000)) {
                    await this.ensureValidToken();
                }
            }
        }, 30 * 1000);
    },

    _persistSubscription() {
        localStorage.setItem("fa_has_subscription", String(AppState.hasSubscription));
        localStorage.setItem("fa_subscription_id", AppState.subscriptionId || "");
        const _planToSave = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null' && AppState.subscriptionPlan !== 'undefined') ? AppState.subscriptionPlan : "";
        localStorage.setItem("fa_subscription_plan", _planToSave);
    },

    /**
     * Clears all auth + subscription state and returns to welcome screen.
     */
    logout() {
        if (this._tokenRefreshInterval) {
            clearInterval(this._tokenRefreshInterval);
            this._tokenRefreshInterval = null;
        }

        ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data on logout: ", err));

        const lastEmail = localStorage.getItem("fa_user_email");
        const lastName = localStorage.getItem("fa_user_name");
        const lastProvider = localStorage.getItem("fa_user_provider");
        if (lastEmail) localStorage.setItem("fa_last_user_email", lastEmail);
        if (lastName) localStorage.setItem("fa_last_user_name", lastName);
        if (lastProvider) localStorage.setItem("fa_last_user_provider", lastProvider);

        const currentJwt = AppState.jwtToken;
        const currentRefreshToken = AppState.refreshToken;

        try {
            if (currentRefreshToken || currentJwt) {
                ApiService.apiFetch("/api/auth/logout", {
                    method: "POST",
                    body: JSON.stringify({ refreshToken: currentRefreshToken, token: currentJwt })
                }).catch(() => { });
            }
        } catch (_) { }

        AppState.userEmail = null;
        AppState.userName = null;
        AppState.userProvider = null;
        AppState.hasSubscription = false;
        AppState.subscriptionId = null;
        AppState.subscriptionPlan = null;
        AppState.erpConnected = false;
        AppState.erpType = null;
        AppState.erpOrgName = null;
        AppState.erpConnectedDate = null;
        AppState.jwtToken = null;  // Clear the JWT token on logout
        AppState.refreshToken = null; // Clear the refresh token on logout

        // Notification history is per-user server-side data — drop the
        // in-memory cache and badge/drawer on logout so nothing from
        // this account is still visible (even for an instant) if a
        // different account signs in next in this same taskpane session.
        NotificationService._cache = [];
        NotificationService._lastNotif = null;
        NotificationService.renderBadge();
        const notifDrawerEl = document.getElementById("notifDrawer");
        if (notifDrawerEl) notifDrawerEl.style.display = "none";

        hideBanner();

        [
            "fa_user_email", "fa_user_name", "fa_user_provider",
            "fa_has_subscription", "fa_subscription_id", "fa_subscription_plan",
            "fa_erp_connected", "fa_erp_type", "fa_erp_org", "fa_erp_date",
            "fa_current_company_id",
            "fa_last_view", "fa_jwt_token", "fa_refresh_token"
        ].forEach(k => localStorage.removeItem(k));

        ViewRouter.show("Welcome");
    }
};

export { AuthService };
