/**
 * Main application controller: startup order, view event binding and session
 * restoration.
 *
 * AppController is assembled from four parts that share one `this`:
 *   - this file             : init(), error view, session restoration
 *   - trialController.js    : free-trial dialog, watcher, expiry modal
 *   - viewBindings.js       : Welcome / Plans / Payment / Success views
 *   - dashboardBindings.js  : the Dashboard view (+ dataActionBindings.js)
 * Merging rather than nesting keeps every `AppController.x()` call site and
 * the original init() ordering intact.
 */
import { AppState } from "../state/appState.js";
import { ViewRouter } from "../router/viewRouter.js";
import { ApiService } from "../services/apiService.js";
import { ExcelService } from "../services/excelService.js";
import { DashboardService } from "../services/dashboardService.js";
import { createTrialController } from "./trialController.js";
import { createViewBindings } from "./viewBindings.js";
import { bindDashboardView } from "./dashboardBindings.js";
import { AuthService } from "../services/authService.js";

const AppController = {


    init() {
        // Note: the QuickBooks/Xero log consoles are populated by
        // DashboardService.renderActiveLogConsole(), called once the
        // active provider/company are resolved during restoreSession()
        // below (via render() -> renderERPSection()) — not here, since
        // AppState.currentCompanyId isn't known yet at this point and
        // the console must be scoped to the correct company from the
        // start, never showing another company's history.
        this.bindTrialExpiredModal();
        this.bindWelcomeView();
        this.bindPlansView();
        this.bindPaymentView();
        this.bindSuccessView();
        this.bindDashboardView();
        this.bindErrorView();
        this.restoreSession();
    },

    ...createTrialController(),
    ...createViewBindings(),
    bindDashboardView,

    // ---- Error View ----
    bindErrorView() {
        document.getElementById("btnRetry")?.addEventListener("click", () => {
            ViewRouter.show("Welcome");
        });
    },

    // ---- Session Restoration ----
    /**
     * On load, checks localStorage to see if the user already has an active session.
     * If so, skips the welcome screen and navigates straight to the dashboard.
     */
    restoreSession() {
        const email = localStorage.getItem("fa_user_email");
        if (email) {
            AppState.userEmail = email;
            AppState.userName = localStorage.getItem("fa_user_name");
            AppState.userProvider = localStorage.getItem("fa_user_provider");
            AppState.hasSubscription = localStorage.getItem("fa_has_subscription") === "true";
            AppState.subscriptionId = localStorage.getItem("fa_subscription_id");
            AppState.subscriptionPlan = (v => (!v || v === 'null' || v === 'undefined') ? null : v)(localStorage.getItem("fa_subscription_plan"));
            AppState.trialEndsAt = (v => (v ? parseInt(v, 10) : null))(localStorage.getItem("fa_trial_ends_at"));
            AppState.erpConnected = localStorage.getItem("fa_erp_connected") === "true";
            AppState.erpType = localStorage.getItem("fa_erp_type");
            AppState.currentCompanyId = localStorage.getItem("fa_current_company_id") || null;

            // Resume/re-check the trial countdown on reload — if the
            // 2 minutes already elapsed while the task pane was closed,
            // this shows the upgrade popup immediately.
            AppController.startTrialExpirationWatcher();

            // Check if we are already locally expired to pop the modal immediately before rendering old UI
            if (AppController.isTrialExpired()) {
                const modal = document.getElementById("trialExpiredModal");
                if (modal && modal.style.display !== "flex") {
                    modal.style.display = "flex";
                    ExcelService.clearMasterData().catch(e => console.error(e));
                }
            }

            DashboardService.render();

            const lastView = localStorage.getItem("fa_last_view");
            if (lastView && lastView !== "Welcome" && lastView !== "Loading" && lastView !== "Error") {
                ViewRouter.show(lastView);
            } else {
                ViewRouter.show("Dashboard");
            }

            // Cached localStorage can go stale (plan changed from
            // another device/session), so use /api/auth/me to pick up
            // a newer confirmed plan. This must only ever CONFIRM or
            // UPGRADE the displayed plan — it must never clear it back
            // to "no plan" (which renders as Basic), since a failed
            // request, an offline moment, or a write that simply
            // hasn't landed yet would otherwise look exactly like the
            // plan silently reverting to Basic on refresh.
            if (AppState.jwtToken) {
                AuthService.startTokenRefreshTimer();
                ApiService.apiFetch("/api/auth/me")
                    .then(r => (r.ok ? r.json() : null))
                    .then(result => {
                        const dbPlan = result?.user?.plan || null;
                        if (dbPlan && dbPlan !== AppState.subscriptionPlan) {
                            AppState.subscriptionPlan = dbPlan;
                            AppState.hasSubscription = true;
                            localStorage.setItem("fa_subscription_plan", dbPlan);
                            localStorage.setItem("fa_has_subscription", "true");
                            DashboardService.render();
                            DashboardService.renderERPSection();
                        }
                        // If dbPlan is empty/missing, do nothing — keep
                        // showing whatever plan was already cached.

                        const dbTrialEndsAt = result?.user?.trialEndsAt ? new Date(result.user.trialEndsAt).getTime() : null;
                        if (dbTrialEndsAt && dbTrialEndsAt !== AppState.trialEndsAt) {
                            AppState.trialEndsAt = dbTrialEndsAt;
                            localStorage.setItem("fa_trial_ends_at", String(dbTrialEndsAt));
                        }

                        // Check expiration explicitly off the backend source of truth
                        if (AppController.isTrialExpired()) {
                            ExcelService.clearMasterData().catch(e => console.error(e));
                            const currentView = ViewRouter.getCurrentView ? ViewRouter.getCurrentView() : (localStorage.getItem("fa_last_view") || "");
                            const isOnUpgradeScreens = currentView === "Plans" || currentView === "Payment" || currentView === "Success";
                            const modal = document.getElementById("trialExpiredModal");
                            if (modal && modal.style.display !== "flex" && !isOnUpgradeScreens) {
                                modal.style.display = "flex";
                            }
                        } else {
                            // If they paid and are now active, ensure modal is hidden
                            const modal = document.getElementById("trialExpiredModal");
                            if (modal) modal.style.display = "none";
                        }

                        // Re-run with whatever fresh plan/expiry we just got —
                        // covers both "just started a real trial" and "trial
                        // already ended while the task pane was closed".
                        AppController.startTrialExpirationWatcher();
                    })
                    .catch(() => {
                        // Offline or request failed — keep showing the
                        // cached plan rather than blocking the UI.
                    });
            }
        } else {
            // Reset AppState to defaults
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

            // Always show welcome screen
            ViewRouter.show("Welcome");
        }
    }
};

export { AppController };
