/**
 * DashboardService — ERP OAuth launch, connection finalisation and disconnect.
 *
 * Mixed into DashboardService (see dashboardService.js). Kept apart from the
 * rendering half because it owns the popup/dialog lifecycle: the redirect
 * transition card, the postMessage handshake and the single-shot
 * cancel/complete guard that has to survive both the popup and the Office
 * dialog fallback.
 */
import { AppState } from "../state/appState.js";
import { ApiService } from "./apiService.js";
import { ExcelService } from "./excelService.js";
import { DashboardService } from "./dashboardService.js";

export function createErpConnection() {
    return {

        /**
         * Launches the ERP OAuth popup for the given provider.
         *
         * @param {"quickbooks"|"xero"} provider
         * @param {string|null} [reconnectId] The company/tenant id the user
         *   explicitly clicked "Reconnect" on. Passed through to the backend
         *   as ?reconnectId so it can refuse the authorization if a
         *   different company is chosen in the provider's own account
         *   picker — a reconnect must restore the company it was started
         *   for, never quietly add a new one. Omitted (null) for a normal
         *   "Add Company" flow.
         */
        launchERPOAuth(provider, reconnectId = null) {
            // Reentrancy guard — the redirect card's overlay blocks clicks
            // underneath it once shown, but this also covers any
            // programmatic re-entry (double keydown, a second call before
            // the first paint) so a second attempt can never stack on top
            // of one already in flight.
            if (AppState.erpAuthInProgress) return;
            AppState.erpAuthInProgress = true;

            this.showConnecting(provider);
            AppState.currentProvider = provider;
            const isQB = provider === "quickbooks";
            const pName = isQB ? "QuickBooks" : "Xero";

            const encodedMail = encodeURIComponent(AppState.userEmail || "");
            const tokenParam = AppState.jwtToken ? `&token=${encodeURIComponent(AppState.jwtToken)}` : "";
            const reconnectParam = reconnectId ? `&reconnectId=${encodeURIComponent(reconnectId)}` : "";
            const connectUrl = isQB
                ? `${ApiService.BASE}/api/quickbooks/connect/?tier=${AppState.currentTier}&mail=${encodedMail}${tokenParam}${reconnectParam}`
                : `${ApiService.BASE}/api/xero/connect?tier=${AppState.currentTier}&mail=${encodedMail}${tokenParam}${reconnectParam}`;

            // Popup opening is not a completed action — only logged, never toasted.
            this.addLog(`Opening ${pName} sign-in...`);

            // Guards against the completion/cancellation handling running
            // twice for the same attempt (e.g. the real "connected" message
            // arrives right as the window-closed poll also fires) — only
            // the first one wins, so exactly one outcome is ever processed
            // per attempt. Also clears the reentrancy guard above, whichever
            // way this attempt ends.
            let settled = false;
            const finishOnce = (fn) => {
                if (settled) return;
                settled = true;
                AppState.erpAuthInProgress = false;
                fn();
            };

            // Tracks whatever popup/dialog is currently open for this
            // attempt so the close (×) button on the redirect card can
            // shut it down too, not just the countdown.
            let activeDialog = null;
            let activePopupWin = null;
            let cancelled = false;

            // Once the sign-in window/dialog successfully opens, swap the
            // card from Phase 1 (5-second countdown) to Phase 2 (a
            // lightweight "Waiting for you to sign in..." spinner) — but
            // only once BOTH are true: the popup is confirmed open, AND
            // the full 5-second countdown animation has played out,
            // whichever finishes second. Without that floor, a dialog
            // that opens quickly (a few hundred ms on some hosts) would
            // cut the "Redirecting..." animation short and look glitchy.
            let countdownDone = false;
            let popupReadyToClose = false;
            const attemptShowWaitingState = () => {
                if (countdownDone && popupReadyToClose) this.showRedirectWaitingState(provider);
            };
            const markPopupOpened = () => {
                popupReadyToClose = true;
                attemptShowWaitingState();
            };
            // If the popup/dialog definitively fails to open (blocked,
            // etc.) there's nothing left to wait for — close the card
            // immediately (no countdown floor) and fall back to the
            // normal "provider selected" screen so the error status
            // shows over a real screen instead of a stuck redirect card.
            const closeRedirectCardOnFailure = () => this.showProviderSelected(provider);

            const openERPPopup = () => {
                if (typeof Office !== "undefined" && Office.context && Office.context.ui) {
                    Office.context.ui.displayDialogAsync(
                        connectUrl,
                        { height: 60, width: 45, displayInIframe: false, promptBeforeOpen: true },
                        (asyncResult) => {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                const win = window.open(connectUrl, "_blank", "width=800,height=600");
                                activePopupWin = win;
                                if (!win) {
                                    AppState.erpAuthInProgress = false;
                                    closeRedirectCardOnFailure();
                                    this.showStatus(`Unable to open ${pName} sign-in. Allow popups.`, "error", null, provider);
                                } else {
                                    markPopupOpened();
                                    const timer = setInterval(() => {
                                        if (win.closed) {
                                            clearInterval(timer);
                                            // Window closed with no completion message received —
                                            // treat as a plain user cancellation: no verification,
                                            // no callback, no toast. Just restore the prior screen.
                                            finishOnce(() => DashboardService.cancelERPConnection(provider));
                                        }
                                    }, 1000);
                                }
                            } else {
                                const dialog = asyncResult.value;
                                activeDialog = dialog;
                                markPopupOpened();
                                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                                    if (arg.message === "qb_connected" || arg.message === "xero_connected") {
                                        dialog.close();
                                        finishOnce(() => DashboardService.onERPConnected(provider));
                                    }
                                });
                                // Fallback: dialog closed manually (error 12006) without a
                                // completion message — this is a user cancellation, not a
                                // failure. No backend verification, no callback, no toast:
                                // just silently restore the screen the user started from.
                                dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                                    if (arg.error === 12006) {
                                        finishOnce(() => DashboardService.cancelERPConnection(provider));
                                    }
                                });
                            }
                        }
                    );
                } else {
                    // Standard popup (browser context)
                    const msgHandler = (event) => {
                        if (event.data === "qb_connected" || event.data === "xero_connected") {
                            window.removeEventListener("message", msgHandler);
                            finishOnce(() => DashboardService.onERPConnected(provider));
                        }
                    };
                    window.addEventListener("message", msgHandler);

                    const win = window.open(connectUrl, `${provider}_auth`, "width=800,height=600");
                    activePopupWin = win;
                    if (!win) {
                        AppState.erpAuthInProgress = false;
                        closeRedirectCardOnFailure();
                        this.showStatus(`Unable to open ${pName} sign-in. Allow popups.`, "error", null, provider);
                    } else {
                        markPopupOpened();
                        // Fallback: window closed manually without a completion
                        // message — this is a user cancellation. No backend
                        // verification, no completion callback, no toast: just
                        // silently restore the screen the user started from.
                        const timer = setInterval(() => {
                            if (win.closed) {
                                clearInterval(timer);
                                window.removeEventListener("message", msgHandler);
                                finishOnce(() => DashboardService.cancelERPConnection(provider));
                            }
                        }, 1000);
                    }
                }
            };

            // Open the OAuth popup/dialog NOW, synchronously, in the same
            // call stack as the click that triggered launchERPOAuth().
            //
            // This used to be deferred behind the 5-second countdown (a
            // setTimeout/setInterval callback firing openERPPopup() after
            // the delay) — but Office.context.ui.displayDialogAsync, and
            // the window.open() it and its browser-context fallback both
            // rely on, only reliably work when triggered directly by a
            // user gesture. Once the call happens inside an async timer
            // callback instead of the click handler itself, the host/
            // browser's popup blocker silently swallows it: the redirect
            // card would finish its countdown ("Opening secure sign-in...",
            // bar full, all dots lit) and then nothing would actually
            // open. Calling it here keeps that gesture intact — the sign-in
            // window/dialog is already loading in the background — while
            // the countdown below still plays out its full 5 seconds as a
            // purely cosmetic "Redirecting..." loading screen in the task
            // pane, which lines up naturally with the popup's own load time.
            openERPPopup();

            // Drive the redirect card's 5-second progress bar / dot
            // sequence / countdown text purely for show — see above, this
            // no longer gates when the popup actually opens.
            // TODO(review): 5000ms is currently hardcoded per the request
            // ("wait for exactly 5 seconds") — flag if this should instead
            // be a configurable value (e.g. sourced from config/AppState).
            const REDIRECT_DELAY_MS = 5000;
            const REDIRECT_TICK_MS = 1000;
            const totalTicks = Math.round(REDIRECT_DELAY_MS / REDIRECT_TICK_MS);
            let secondsLeft = totalTicks;

            const progressFillEl = document.getElementById("redirectProgressFill");
            const countdownEl = document.getElementById("redirectCountdown");
            const dotEls = document.querySelectorAll("#redirectDots .dot");

            const redirectTimer = setInterval(() => {
                secondsLeft -= 1;
                const elapsedTicks = totalTicks - secondsLeft;
                if (progressFillEl) {
                    progressFillEl.style.width = `${Math.min(100, Math.round((elapsedTicks / totalTicks) * 100))}%`;
                }
                dotEls.forEach((dot, i) => dot.classList.toggle("active", i < elapsedTicks));
                if (countdownEl) {
                    countdownEl.textContent = secondsLeft > 0
                        ? `Please wait, opening in ${secondsLeft} second${secondsLeft === 1 ? "" : "s"}...`
                        : "Opening secure sign-in...";
                }
                if (secondsLeft <= 0) {
                    clearInterval(redirectTimer);
                    countdownDone = true;
                    attemptShowWaitingState();
                }
            }, REDIRECT_TICK_MS);

            // Close (×) button on the redirect card — lets the user back
            // out during the 5s hold, or while the popup itself is open,
            // instead of being stuck waiting. The listener is bound once
            // and always defers to whichever attempt is currently active.
            const closeBtn = document.getElementById("redirectCloseBtn");
            if (closeBtn && !closeBtn.dataset.bound) {
                closeBtn.dataset.bound = "true";
                closeBtn.addEventListener("click", () => {
                    if (typeof DashboardService._activeRedirectCancel === "function") {
                        DashboardService._activeRedirectCancel();
                    }
                });
            }
            DashboardService._activeRedirectCancel = () => {
                if (cancelled) return;
                cancelled = true;
                clearInterval(redirectTimer);
                try {
                    if (activeDialog && typeof activeDialog.close === "function") activeDialog.close();
                } catch (_) { /* dialog already gone — nothing to clean up */ }
                try {
                    if (activePopupWin && !activePopupWin.closed) activePopupWin.close();
                } catch (_) { /* popup already gone — nothing to clean up */ }
                finishOnce(() => DashboardService.cancelERPConnection(provider));
            };
        },

        /**
         * Called when the user manually closes the OAuth popup/dialog
         * without completing authentication. This is a pure cancellation:
         * no backend verification, no onERPConnected callback, and no
         * success/error/warning toast of any kind — closing the popup is
         * not an outcome that gets reported to the user. It restores the
         * dashboard to its real state (the disconnected provider-choice
         * screen, since no connection exists) rather than leaving the
         * "Not connected" provider-selected screen with its Setup/Pull
         * buttons visible — that intermediate screen only makes sense
         * once a connection attempt is in flight, not after it's been
         * abandoned. Only a genuine OAuth callback from the backend
         * (handled by onERPConnected) ever proceeds to verification and
         * notifications.
         * @param {"quickbooks"|"xero"} provider
         */
        cancelERPConnection(provider) {
            void provider; // kept for signature symmetry with the completion path
            this.renderERPSection();
        },

        /**
         * Handles successful ERP OAuth callback — saves state and refreshes dashboard.
         * @param {"quickbooks"|"xero"} provider
         */
        onERPConnected(provider) {
            const isQB = provider === "quickbooks";
            const now = new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" });

            AppState.erpConnected = true;
            AppState.erpType = provider;
            AppState.erpConnectedDate = now;
            AppState.currentProvider = provider;

            localStorage.setItem("fa_erp_connected", "true");
            localStorage.setItem("fa_erp_type", provider);
            localStorage.setItem("fa_erp_date", now);

            // Attempt to fetch org name from backend. Uses apiFetch (not raw
            // fetch) so the JWT is actually attached — this endpoint is
            // authenticated, so without it every call here 401'd
            // unconditionally and silently fell back to a random ID below.
            const tokenPath = isQB ? "/api/quickbooks/tokens/" : "/api/xero/tokens";

            ApiService.apiFetch(tokenPath)
                .then(r => (r.ok ? r.json() : { tokens: [] }))
                .then(data => {
                    const tokens = data.tokens || [];
                    const realmId = tokens[0]?.realm_id || tokens[0]?.tenant_name || null;
                    // Use backend realm_id if available, else generate a random 16-digit ID
                    const connId = realmId || DashboardService._generateConnectionId();
                    AppState.erpOrgName = connId;
                    AppState.connectionId = connId;
                    localStorage.setItem("fa_erp_org", connId);
                    DashboardService._finalizeConnection(provider, connId);
                })
                .catch(() => {
                    // Backend not available — generate a random connection ID
                    const connId = DashboardService._generateConnectionId();
                    AppState.erpOrgName = connId;
                    AppState.connectionId = connId;
                    localStorage.setItem("fa_erp_org", connId);
                    DashboardService._finalizeConnection(provider, connId);
                });
        },

        /**
         * Generates a random 16-digit numeric connection/realm ID.
         * @returns {string}
         */
        _generateConnectionId() {
            // Generate 16-digit number similar to QuickBooks realm IDs
            const part1 = Math.floor(1000000000 + Math.random() * 9000000000); // 10 digits
            const part2 = Math.floor(100000 + Math.random() * 900000);          // 6 digits
            return String(part1) + String(part2);
        },

        /**
         * Finalises connection: marks step 1 complete, renders connected dashboard, updates ID.
         * @param {string} provider
         * @param {string} connId
         */
        _finalizeConnection(provider, connId) {
            const isQB = provider === "quickbooks";

            // Transition to fully connected dashboard (Image 3)
            this.render();

            // Explicitly set the realm ID text in connected header
            const realmEl = document.getElementById("connRealmId");
            if (realmEl) realmEl.textContent = connId;

            // Show success status
            this.showStatus(`${isQB ? "QuickBooks" : "Xero"} connected successfully.`, "success", null, provider);

            // Step 1 (Connect) is derived live from AppState.erpConnected
            // (already true by this point) — this applies it to both the
            // provider-selected and connected step indicators.
            this.applyStepState();
        },

        /**
         * Disconnects the ERP provider — clears state but keeps FinAccrual subscription.
         */
        async disconnectERP() {
            // Optimistically update AppState
            AppState.erpConnected = false;
            AppState.erpType = null;
            AppState.erpOrgName = null;
            AppState.erpConnectedDate = null;
            AppState.isConnected = false;
            AppState.connectionId = null;
            AppState.currentCompanyId = null;
            AppState.forceWelcome = true;

            localStorage.removeItem("fa_erp_connected");
            localStorage.removeItem("fa_erp_type");
            localStorage.removeItem("fa_erp_org");
            localStorage.removeItem("fa_erp_date");

            // resetSteps() clears every step indicator (provider-selected
            // and connected consoles, both providers) — and, since this is
            // an explicit full disconnect (not a refresh), also wipe the
            // persisted Setup/Pull completion state behind them for real.
            this.resetSteps();
            this._saveStepState({});

            // Clear logs — the console is in-memory/session-only already
            // (see DashboardService._sessionLogs above), this just also
            // wipes what accumulated earlier in this same session.
            const qbLog = document.getElementById("qbLog");
            const xeroLog = document.getElementById("xeroLog");
            if (qbLog) qbLog.innerHTML = "";
            if (xeroLog) xeroLog.innerHTML = "";
            this._saveLogs([]);

            this.showStatus("Disconnecting all companies...", "success");

            // Disconnect ALL companies for this user from the backend
            try {
                const mail = AppState.userEmail || "";
                const connsRes = await ApiService.apiFetch(`/api/connections?mail=${encodeURIComponent(mail)}`);
                const conns = await connsRes.json();
                const activeConns = (conns || []).filter(c => c.status !== 'Disconnected');
                await Promise.all(activeConns.map(c =>
                    ApiService.apiFetch(`/api/connections/${c.companyId}`, { method: "DELETE" }).catch(() => { })
                ));
            } catch (_) { }

            try { await ExcelService.clearMasterData(); } catch (_) { }

            // Now re-render — all companies should be Disconnected, so it shows the disconnected view
            this.renderERPSection();
            this.showStatus("QuickBooks disconnected. Your FinAccrual account is still active.", "success");
        }
    };
}
