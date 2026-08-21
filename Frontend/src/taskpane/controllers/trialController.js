/**
 * Free-trial lifecycle: the trial-select dialog, the expiry watcher and the
 * "trial expired" modal.
 *
 * Mixed into AppController (see appController.js) so `AppController.isOnTrial()`
 * and friends keep resolving exactly as before. Exported as a hoisted factory
 * so the mix-in survives the controller <-> service import cycle.
 */
import { AppState } from "../state/appState.js";
import { ViewRouter } from "../router/viewRouter.js";
import { ApiService } from "../services/apiService.js";
import { ExcelService } from "../services/excelService.js";
import { AuthService } from "../services/authService.js";
import { DashboardService } from "../services/dashboardService.js";
import { AppController } from "./appController.js";

export function createTrialController() {
    return {

        // ---- Trial Select Dialog ----
        openTrialSelectDialog() {
            const dialogUrl = window.location.origin + "/trialselect.html";
            let dialog = null;

            Office.context.ui.displayDialogAsync(dialogUrl, { height: 60, width: 45, displayInIframe: true }, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to open trial select dialog:", asyncResult.error.message);
                } else {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                        try {
                            const message = JSON.parse(arg.message);
                            if (message.type === 'START_TRIAL') {
                                dialog.close();
                                // Ask the backend to actually start the trial
                                // (sets plan + a real trial_ends_at clock
                                // starting now) rather than faking it
                                // client-side, so the 2-minute countdown and
                                // the 1-company limit are both enforced from
                                // a real server timestamp.
                                ApiService.startTrial()
                                    .then((result) => {
                                        const user = result?.user || {};
                                        AppState.hasSubscription = true;
                                        AppState.subscriptionId = user.subscriptionId || AppState.subscriptionId;
                                        AppState.subscriptionPlan = user.plan || "trial";
                                        AppState.trialEndsAt = user.trialEndsAt ? new Date(user.trialEndsAt).getTime() : null;
                                        if (AppState.trialEndsAt) {
                                            localStorage.setItem("fa_trial_ends_at", String(AppState.trialEndsAt));
                                        }
                                        AuthService._persistSubscription();
                                        DashboardService.render();
                                        ViewRouter.show("Dashboard");
                                        DashboardService.showStatus("Free Trial started successfully!", "success");
                                        // Legacy local-timer fallback, kept in sync in case
                                        // trialEndsAt couldn't be read from the response.
                                        localStorage.setItem("fa_trial_start", Date.now().toString());
                                        AppController.startTrialExpirationWatcher();
                                    })
                                    .catch((err) => {
                                        console.error("Failed to start free trial:", err);
                                        DashboardService.showStatus("Couldn't start your free trial. Please try again.", "error");
                                    });
                            } else if (message.type === 'VIEW_PLANS') {
                                dialog.close();
                                ViewRouter.show("Plans");
                            }
                        } catch (e) {
                            console.error("Error parsing dialog message:", e);
                        }
                    });
                }
            });
        },

        // ---- Trial Expiration Logic ----
        // Fallback only, used for the client-only mock trial started from
        // trialselect.html (no backend call, so no server trialEndsAt
        // exists for it). Real accounts get a server-issued trial_ends_at
        // (2 minutes by default — see Backend core/config TRIAL.DURATION_MS)
        // which is authoritative whenever it's available.
        TRIAL_DURATION_MS: 2 * 60 * 1000,
        _trialWatcherId: null,

        isOnTrial() {
            return (AppState.subscriptionPlan || "").toLowerCase().includes("trial");
        },

        // Resolves the ms-epoch timestamp the trial ends at, preferring the
        // real backend value (AppState.trialEndsAt, from /api/auth/me) over
        // the local mock-flow timer.
        getTrialEndTimestamp() {
            if (AppState.trialEndsAt) return AppState.trialEndsAt;
            const trialStartStr = localStorage.getItem("fa_trial_start");
            if (!trialStartStr) return null;
            return parseInt(trialStartStr, 10) + AppController.TRIAL_DURATION_MS;
        },

        checkTrialExpiration() {
            if (!AppController.isOnTrial()) {
                // Plan changed (e.g. user upgraded) — no need to keep polling.
                AppController.stopTrialExpirationWatcher();
                return;
            }

            const endTs = AppController.getTrialEndTimestamp();
            if (!endTs) return;

            if (Date.now() >= endTs) {
                // Trial local time is up. Stop watcher and check real backend status
                AppController.stopTrialExpirationWatcher();

                ApiService.apiFetch("/api/auth/me")
                    .then(r => (r.ok ? r.json() : null))
                    .then(result => {
                        if (result && result.user) {
                            const u = result.user;
                            const currentPlan = (u.plan || "").toLowerCase();
                            const actualEndTs = u.trialEndsAt ? new Date(u.trialEndsAt).getTime() : null;
                            const isExpired = currentPlan === 'expired' || (currentPlan.includes('trial') && actualEndTs && Date.now() >= actualEndTs);

                            if (isExpired) {
                                AppState.subscriptionPlan = u.plan;
                                localStorage.setItem("fa_subscription_plan", u.plan);

                                const modal = document.getElementById("trialExpiredModal");
                                if (modal && modal.style.display !== "flex") {
                                    modal.style.display = "flex";
                                    ExcelService.clearMasterData().catch(e => console.error("Failed to clear master data on trial expiry", e));
                                }
                            } else {
                                // Backend confirms still active (e.g. upgraded on another device)
                                AppState.subscriptionPlan = u.plan;
                                localStorage.setItem("fa_subscription_plan", u.plan);
                                if (actualEndTs) {
                                    AppState.trialEndsAt = actualEndTs;
                                    localStorage.setItem("fa_trial_ends_at", String(actualEndTs));
                                }
                                DashboardService.render();
                                AppController.startTrialExpirationWatcher();
                            }
                        } else {
                            throw new Error("No user in response");
                        }
                    })
                    .catch(e => {
                        console.error("Failed to check subscription status on trial expiry", e);
                        const modal = document.getElementById("trialExpiredModal");
                        if (modal && modal.style.display !== "flex") {
                            modal.style.display = "flex";
                            ExcelService.clearMasterData().catch(err => console.error(err));
                        }
                    });
            }
        },

        // Polls once a second so the "Upgrade Now" popup appears the moment
        // the trial window elapses, instead of only being checked one time
        // right when the trial starts (which always read as 0ms elapsed and
        // so never actually showed the popup).
        startTrialExpirationWatcher() {
            AppController.stopTrialExpirationWatcher();
            if (!AppController.isOnTrial()) return;

            // If we don't have the server-issued expiry yet for this trial
            // (e.g. right after a fresh login), fetch it once so the popup
            // is timed off the real 2-minute clock rather than just the
            // local mock-flow fallback.
            if (!AppState.trialEndsAt && AppState.jwtToken) {
                ApiService.apiFetch("/api/auth/me")
                    .then(r => (r.ok ? r.json() : null))
                    .then(result => {
                        const ts = result?.user?.trialEndsAt ? new Date(result.user.trialEndsAt).getTime() : null;
                        if (ts) {
                            AppState.trialEndsAt = ts;
                            localStorage.setItem("fa_trial_ends_at", String(ts));
                            AppController.checkTrialExpiration();
                        }
                    })
                    .catch(() => { /* fall back to the local timer, if any */ });
            }

            // Check immediately in case the trial already expired (e.g. the
            // task pane was reopened after the window elapsed), then keep
            // polling so it fires the instant the trial ends.
            AppController.checkTrialExpiration();
            AppController._trialWatcherId = setInterval(() => {
                AppController.checkTrialExpiration();
            }, 1000);
        },

        stopTrialExpirationWatcher() {
            if (AppController._trialWatcherId) {
                clearInterval(AppController._trialWatcherId);
                AppController._trialWatcherId = null;
            }
        },

        // ---- Trial Expired Modal ----
        bindTrialExpiredModal() {
            const modal = document.getElementById("trialExpiredModal");
            const btnClose = document.getElementById("btnCloseTrialExpired");
            const btnUpgrade = document.getElementById("btnUpgradeNow");

            if (btnClose) {
                btnClose.addEventListener("click", () => {
                    if (modal) modal.style.display = "none";
                });
            }

            if (btnUpgrade) {
                btnUpgrade.addEventListener("click", () => {
                    if (modal) modal.style.display = "none";
                    ViewRouter.show("Plans");
                });
            }
        },
    };
}
