/**
 * DashboardService — activity log console and progress-step state.
 *
 * Mixed into DashboardService (see dashboardService.js) so every call site
 * stays `DashboardService.addLog(...)` / `DashboardService.markStepComplete(...)`.
 * Exported as a hoisted factory so the mix-in is safe against the import
 * cycle between the dashboard modules and the services they call back into.
 */
import { AppState } from "../state/appState.js";
import { ViewRouter } from "../router/viewRouter.js";
import { NotificationService } from "./notificationService.js";

export function createDashboardConsole() {
    return {

        /**
         * Updates the progress step markers for the ERP console.
         */
        renderERPConsole() {
            const stepId = AppState.erpType === "quickbooks" ? "stepConnect" : "xeroStepConnect";
            const stepEl = document.getElementById(stepId);
            if (stepEl) stepEl.classList.add("complete");
        },

        // Console log history is in-memory ONLY for the current taskpane
        // session — it is never written to or restored from localStorage.
        // A fresh taskpane load (reopening Excel, reconnecting, switching
        // back to a company) always starts with an empty console: an entry
        // can only ever appear after the real action it describes has
        // actually run in THIS session, never carried over from an earlier
        // one. Capped so it can't grow unbounded during a long session.
        //
        // Every entry is still tagged with both the provider AND the
        // company that was active when it was logged, and the console only
        // ever renders entries matching the CURRENT provider + active
        // company — that's what stops a log line from a different company
        // (logged earlier in this same session) from showing up as if it
        // just happened here.
        MAX_LOG_ENTRIES: 300,
        _sessionLogs: [],

        /** @returns {Array<{provider:('quickbooks'|'xero'), companyId:(string|null), message:string, timestamp:string}>} */
        _loadLogs() {
            return this._sessionLogs;
        },

        _saveLogs(list) {
            this._sessionLogs = list;
        },

        /**
         * Renders a single stored/new entry as a log line and appends it to
         * the given console element, without re-stamping the time — the
         * original timestamp is preserved exactly as logged.
         * @param {HTMLElement} log
         * @param {string} message
         * @param {string} timestampIso
         */
        _appendLogLine(log, message, timestampIso) {
            const line = document.createElement("div");
            line.className = "log-line";
            if (message.toLowerCase().includes("error")) {
                line.style.color = "#ef4444"; // Red color for errors
            }
            const timeLabel = new Date(timestampIso).toLocaleTimeString();
            line.textContent = `[${timeLabel}] ${message}`;
            log.appendChild(line);
        },

        /**
         * Re-renders the visible QuickBooks/Xero console from persisted
         * history, filtered strictly to the current provider AND the
         * currently active company (AppState.currentCompanyId). Call this
         * any time the active provider or company changes (connect,
         * switch, resume, disconnect, dropdown/modal change) as well as on
         * app init — it's the single source of truth for what the console
         * shows, so a company that's never had Setup/Pull run always
         * starts with a clean console, never another company's history.
         */
        renderActiveLogConsole() {
            const provider = AppState.currentProvider === "quickbooks" ? "quickbooks" : "xero";
            const logId = provider === "quickbooks" ? "qbLog" : "xeroLog";
            const log = document.getElementById(logId);
            if (!log) return;

            const companyId = AppState.currentCompanyId || null;
            const entries = this._loadLogs().filter(entry =>
                entry.provider === provider && (entry.companyId || null) === companyId
            );

            log.innerHTML = "";
            entries.forEach(entry => this._appendLogLine(log, entry.message, entry.timestamp));
            log.scrollTop = log.scrollHeight;
        },

        /**
         * Adds a log line for an action that has actually just executed,
         * persists it tagged to the current provider + active company, and
         * re-renders the console from that persisted history. Because the
         * console always re-derives its content from storage (filtered to
         * the exact provider/company in view), a log entry can only ever
         * appear after the real action it describes has run — nothing is
         * fabricated or carried over from a different company.
         * @param {string} message
         */
        addLog(message) {
            const provider = AppState.currentProvider === "quickbooks" ? "quickbooks" : "xero";
            const companyId = AppState.currentCompanyId || null;
            const timestamp = new Date().toISOString();

            const list = this._loadLogs();
            list.push({ provider, companyId, message: String(message), timestamp });
            if (list.length > this.MAX_LOG_ENTRIES) {
                list.splice(0, list.length - this.MAX_LOG_ENTRIES);
            }
            this._saveLogs(list);

            this.renderActiveLogConsole();
        },

        /**
         * Pops an immediate top-right toast and records the same
         * notification in the Bell Notification Center. Transient
         * "in progress" messages (e.g. "Pulling data...") are not turned
         * into a toast/notification — only terminal outcomes are, so the
         * user isn't shown a toast for every intermediate step. No banner
         * is rendered inline; status updates live in the log console only.
         *
         * `provider` tags a QuickBooks/Xero-specific outcome so it's only
         * ever shown while the user is actively on that same ERP — pass
         * "quickbooks" or "xero" for provider-scoped actions (connect,
         * setup sheets, pull data, disconnect a company, etc.). Leave it
         * out for actions that aren't tied to either ERP (login, payment,
         * logout, disconnect-everything).
         * @param {string} message
         * @param {"success"|"error"} type
         * @param {string} [detail] - optional second line shown under the toast title and in the bell drawer
         * @param {"quickbooks"|"xero"} [provider] - which ERP this belongs to, if any
         */
        showStatus(message, type, detail, provider) {
            if (typeof message === "string" && !message.trim().endsWith("...")) {
                NotificationService.add(message, type, detail, provider);
            }
        },

        // Setup/Pull step completion (the green checkmarks on steps 2 and 3)
        // persists per provider + company, same reasoning as the log
        // console: without this, a taskpane refresh would reset every step
        // indicator to "not done" even though the log clearly shows Setup
        // and Pull already ran — the two must stay consistent. Step 1
        // (Connect) isn't stored here at all; it's derived live from
        // AppState.erpConnected, which is already persisted separately.
        STEP_STORAGE_KEY: "fa_step_state",

        /** @returns {Object<string, {setup?:boolean, pull?:boolean}>} */
        _loadStepState() {
            try {
                const raw = localStorage.getItem(this.STEP_STORAGE_KEY);
                const parsed = raw ? JSON.parse(raw) : {};
                return (parsed && typeof parsed === "object" && !Array.isArray(parsed)) ? parsed : {};
            } catch (_) {
                return {};
            }
        },

        _saveStepState(state) {
            try {
                localStorage.setItem(this.STEP_STORAGE_KEY, JSON.stringify(state));
            } catch (_) {
                // Storage full/unavailable — step completion just won't
                // persist across a refresh, but the app keeps working.
            }
        },

        _stepStateKey(provider, companyId) {
            return `${provider}::${companyId || "none"}`;
        },

        /**
         * Applies step-completion state to every step indicator for the
         * current provider + active company — both the "provider-selected"
         * (provStepX) and "connected" (stepX/xeroStepX) sets of elements.
         * Connect is derived live from AppState.erpConnected; Setup/Pull
         * come from persisted per-company storage. Explicitly clears the
         * "complete" class when a step is NOT done for this company, so
         * switching to (or refreshing on) a company that's never had
         * Setup/Pull run never shows a stale green checkmark left over
         * from a previously active company.
         */
        applyStepState() {
            const isQB = AppState.currentProvider === "quickbooks";
            const provider = isQB ? "quickbooks" : "xero";
            const key = this._stepStateKey(provider, AppState.currentCompanyId);
            const state = this._loadStepState()[key] || {};
            const connectDone = !!AppState.erpConnected;

            [
                [isQB ? "stepConnect" : "xeroStepConnect", connectDone],
                [isQB ? "stepSetup" : "xeroStepSetup", !!state.setup],
                [isQB ? "stepPull" : "xeroStepPull", !!state.pull],
                ["provStepConnect", connectDone],
                ["provStepSetup", !!state.setup],
                ["provStepPull", !!state.pull]
            ].forEach(([id, done]) => {
                document.getElementById(id)?.classList.toggle("complete", done);
            });
        },

        /**
         * Marks Setup or Pull complete for the current provider + active
         * company — persists it so it survives a refresh, then re-applies
         * step state to the DOM immediately. "connect" is a no-op here
         * since it's derived live from AppState.erpConnected instead.
         * @param {"connect"|"setup"|"pull"} stepName
         */
        markStepComplete(stepName) {
            if (stepName === "setup" || stepName === "pull") {
                const provider = AppState.currentProvider === "quickbooks" ? "quickbooks" : "xero";
                const key = this._stepStateKey(provider, AppState.currentCompanyId);
                const state = this._loadStepState();
                state[key] = { ...(state[key] || {}), [stepName]: true };
                this._saveStepState(state);
            }
            this.applyStepState();
        },

        /**
         * Inverse of markStepComplete — clears a persisted step flag so
         * the progress indicator stops showing it as done. Used when Pull
         * Master Data restarts a cycle from scratch: the previous cycle's
         * "pull complete" tick no longer describes the sheet, which is
         * back to holding only the first batch.
         * @param {"connect"|"setup"|"pull"} stepName
         */
        markStepIncomplete(stepName) {
            if (stepName === "setup" || stepName === "pull") {
                const provider = AppState.currentProvider === "quickbooks" ? "quickbooks" : "xero";
                const key = this._stepStateKey(provider, AppState.currentCompanyId);
                const state = this._loadStepState();
                if (state[key]) {
                    state[key] = { ...state[key] };
                    delete state[key][stepName];
                    this._saveStepState(state);
                }
            }
            this.applyStepState();
        },

        /**
         * Checks if a given step ("setup" or "pull") is marked complete for the
         * current provider + active company.
         * @param {"setup"|"pull"} stepName
         * @returns {boolean}
         */
        isStepComplete(stepName) {
            const provider = AppState.currentProvider === "quickbooks" ? "quickbooks" : "xero";
            const key = this._stepStateKey(provider, AppState.currentCompanyId);
            const state = this._loadStepState()[key] || {};
            return !!state[stepName];
        },

        /**
         * Marks a progress step as complete (connected-dashboard console).
         * @param {string} step - base step ID ("stepConnect"|"stepSetup"|"stepPull")
         */
        completeStep(step) {
            const stepName = step === "stepSetup" ? "setup" : step === "stepPull" ? "pull" : "connect";
            this.markStepComplete(stepName);
        },

        resetSteps() {
            ["stepConnect", "stepSetup", "stepPull", "xeroStepConnect", "xeroStepSetup", "xeroStepPull",
                "provStepConnect", "provStepSetup", "provStepPull"]
                .forEach(id => document.getElementById(id)?.classList.remove("complete", "active"));
        },

        /**
         * Shows the error view with a given message.
         * @param {string} message
         */
        showError(message) {
            const msgEl = document.getElementById("errorMessage");
            if (msgEl) msgEl.textContent = message;
            ViewRouter.show("Error");
            // Also surface as an immediate toast + bell entry, same as any
            // other action failure — the full-screen Error view is shown
            // too, but the user shouldn't have to rely on that alone.
            NotificationService.add(message, "error");
        },
    };
}
