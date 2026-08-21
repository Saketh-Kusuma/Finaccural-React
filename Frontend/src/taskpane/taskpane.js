/**
 * FinAccrual Excel Add-in — Task Pane entry point.
 *
 * This file is the orchestrator only: it wires the modules together and boots
 * the application. All behaviour lives in the modules below.
 *
 * Module map
 * ----------
 *  state/appState.js
 *      AppState — the single reactive state object (auth, subscription/plan,
 *      ERP connection, JWT + refresh token, active company). Mirrored to
 *      localStorage so a task-pane reload restores the last session.
 *      Also owns getMaxCompaniesForPlan(), the client-side mirror of the
 *      backend PLAN_LIMITS table.
 *
 *  router/viewRouter.js
 *      ViewRouter — shows one `.view` section at a time, remembers the last
 *      view, and opens/closes the modal overlays.
 *
 *  services/apiService.js
 *      ApiService — every backend call. apiFetch() attaches the JWT, refreshes
 *      an expired access token (queueing concurrent 401s behind one refresh)
 *      and turns failures into ApiError instances with a stable `.code`.
 *
 *  services/authService.js
 *      AuthService — Google / Microsoft OAuth popups, new-user vs returning-user
 *      hand-off, subscription persistence, logout.
 *
 *  services/checkoutService.js
 *      CheckoutService — plan checkout popup and payment verification.
 *
 *  services/excelService.js  (+ services/excelDataMappers.js)
 *      ExcelService — workbook/worksheet/range operations, master-data writes
 *      and formatting. excelDataMappers.js holds the pure record → row helpers
 *      shared with the pull/refresh handlers.
 *
 *  services/dashboardService.js
 *      DashboardService — dashboard rendering, ERP connection status and
 *      company management. Merged at load time with:
 *        services/dashboardConsole.js      — activity log + progress steps
 *        services/erpConnectionService.js  — ERP OAuth launch / connect / disconnect
 *      so that every `DashboardService.x()` call site resolves against one object.
 *
 *  services/notificationService.js
 *      NotificationService — bell badge, drawer history and toasts, scoped to
 *      the active provider + company.
 *
 *  controllers/appController.js
 *      AppController — startup order, error view and session restoration.
 *      Merged at load time with:
 *        controllers/trialController.js     — free-trial dialog, expiry watcher
 *        controllers/viewBindings.js        — Welcome / Plans / Payment / Success
 *        controllers/dashboardBindings.js   — Dashboard view
 *          controllers/dataActionBindings.js — Setup / Pull / Refresh handlers
 *
 * Import order below is deliberate: the leaf modules (state, router) are
 * evaluated first, then the services, then the controllers, so the composed
 * DashboardService and AppController objects are always assembled after the
 * parts they merge in. The service modules reference each other in cycles
 * (ApiService <-> AuthService <-> DashboardService, and every controller back
 * into the services), which is safe because every one of those references is
 * resolved when a method runs, never while a module is still evaluating.
 */

import "./state/appState.js";
import "./router/viewRouter.js";

import "./services/excelDataMappers.js";
import "./services/excelService.js";
import "./services/notificationService.js";
import "./services/apiService.js";
import "./services/dashboardService.js";
import "./services/authService.js";
import "./services/checkoutService.js";

import { AppController } from "./controllers/appController.js";

/**
 * Boot.
 *
 * Office.onReady fires once the host has finished loading the task pane; only
 * then is it safe to touch Office.* or Excel.*, so nothing is bound before it.
 * AppController.init() performs the whole startup sequence in its original
 * order: bind the trial-expired modal, then the Welcome, Plans, Payment,
 * Success, Dashboard and Error views, then restore the stored session — which
 * is what decides whether the user lands on Welcome or straight on Dashboard.
 *
 * The QuickBooks/Xero log consoles are intentionally NOT populated here; they
 * are filled by DashboardService.renderActiveLogConsole() once restoreSession()
 * has resolved the active provider and company, so the console is never shown
 * scoped to the wrong company.
 */
Office.onReady(() => {
    AppController.init();
});
