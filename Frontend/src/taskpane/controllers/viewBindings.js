/**
 * Event binding for the Welcome, Plans, Payment and Success views.
 *
 * Mixed into AppController (see appController.js); AppController.init() calls
 * each bind* method once, in the original order.
 */
import { AppState } from "../state/appState.js";
import { ViewRouter } from "../router/viewRouter.js";
import { AuthService } from "../services/authService.js";
import { CheckoutService } from "../services/checkoutService.js";
import { DashboardService } from "../services/dashboardService.js";
import { AppController } from "./appController.js";

export function createViewBindings() {
    return {

        // ---- Welcome View ----
        bindWelcomeView() {
            let dialog = null;

            // Bind the real Google and Microsoft login buttons
            document.getElementById("btnSignInGoogle")?.addEventListener("click", () => {
                const btn = document.getElementById("btnSignInGoogle");
                if (btn) btn.disabled = true;
                AuthService.openGooglePopup();
                setTimeout(() => { if (btn) btn.disabled = false; }, 3000);
            });

            document.getElementById("btnSignInMicrosoft")?.addEventListener("click", () => {
                const btn = document.getElementById("btnSignInMicrosoft");
                if (btn) btn.disabled = true;
                AuthService.openMicrosoftPopup();
                setTimeout(() => { if (btn) btn.disabled = false; }, 3000);
            });

            // Bind the main FinAccrual Sign In button to open the custom dialog
            document.getElementById("btnSignIn")?.addEventListener("click", () => {
                const btn = document.getElementById("btnSignIn");
                if (btn) btn.disabled = true;

                const savedName = localStorage.getItem("fa_user_name") || localStorage.getItem("fa_last_user_name") || "";
                const savedEmail = localStorage.getItem("fa_user_email") || localStorage.getItem("fa_last_user_email") || "";
                const rawAccounts = localStorage.getItem("fa_accounts_history") || "[]";
                const dialogUrl = window.location.origin + `/accountpicker.html?name=${encodeURIComponent(savedName)}&email=${encodeURIComponent(savedEmail)}&accounts=${encodeURIComponent(rawAccounts)}`;

                Office.context.ui.displayDialogAsync(dialogUrl, { height: 50, width: 35, displayInIframe: true }, (asyncResult) => {
                    setTimeout(() => { if (btn) btn.disabled = false; }, 3000);

                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Failed to open dialog:", asyncResult.error.message);
                    } else {
                        dialog = asyncResult.value;
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                            try {
                                const message = JSON.parse(arg.message);
                                if (message.type === 'USE_ANOTHER') {
                                    dialog.close();
                                    // Hide the primary button, show the Google/Microsoft buttons in the taskpane
                                    const primaryContainer = document.getElementById("primaryAuthContainer");
                                    const secondaryContainer = document.getElementById("secondaryAuthContainer");
                                    if (primaryContainer) primaryContainer.style.display = "none";
                                    if (secondaryContainer) secondaryContainer.style.display = "flex";
                                } else if (message.type === 'USE_EXISTING') {
                                    dialog.close();
                                    const targetEmail = message.email || savedEmail;
                                    const provider = message.provider || localStorage.getItem("fa_user_provider") || "google";
                                    
                                    const currentEmail = localStorage.getItem("fa_user_email");
                                    const currentToken = localStorage.getItem("fa_jwt_token");

                                    if (targetEmail && targetEmail === currentEmail && currentToken) {
                                        AppController.handleReturningUser(targetEmail, message.name || localStorage.getItem("fa_user_name"), provider, currentToken);
                                    } else {
                                        if (provider === "microsoft") {
                                            AuthService.openMicrosoftPopup(targetEmail || undefined);
                                        } else {
                                            AuthService.openGooglePopup(targetEmail || undefined);
                                        }
                                    }
                                }
                            } catch (e) {
                                console.error("Error parsing dialog message:", e);
                            }
                        });
                    }
                });
            });
        },

        // ---- Plans View ----
        bindPlansView() {
            const toggle = document.getElementById("billingCycleToggle");
            const monthLabel = document.getElementById("labelMonthly");
            const yearLabel = document.getElementById("labelYearly");

            const updatePrices = (isYearly) => {
                document.querySelectorAll("[data-monthly][data-yearly]").forEach(btn => {
                    const monthly = parseInt(btn.dataset.monthly);
                    const yearly = parseInt(btn.dataset.yearly);
                    let amountId = "proAmount";
                    if (btn.id === "btnSelectBasic") amountId = "basicAmount";
                    else if (btn.id === "btnSelectStandard") amountId = "standardAmount";

                    const el = document.getElementById(amountId);
                    if (el) el.textContent = isYearly ? yearly : monthly;

                    btn.dataset.activePrice = String(isYearly ? yearly : monthly);
                    btn.dataset.activeCycle = isYearly ? "Yearly" : "Monthly";
                });
                if (monthLabel) monthLabel.classList.toggle("active-label", !isYearly);
                if (yearLabel) yearLabel.classList.toggle("active-label", isYearly);
            };

            // Initialize labels
            if (monthLabel) monthLabel.classList.add("active-label");

            if (toggle) {
                toggle.addEventListener("change", () => updatePrices(toggle.checked));
            }

            // Plan select buttons
            document.querySelectorAll("[data-plan]").forEach(btn => {
                btn.addEventListener("click", () => {
                    const plan = btn.dataset.plan;
                    const cycle = btn.dataset.activeCycle || "Monthly";
                    const price = btn.dataset.activePrice || btn.dataset.monthly || "Custom";

                    if (plan === "Enterprise") {
                        DashboardService.showStatus("Enterprise enquiry sent! Our sales team will contact you.", "success");
                        return;
                    }

                    // Populate payment view
                    const payPlanEl = document.getElementById("paymentPlanName");
                    const payCycleEl = document.getElementById("paymentBillingCycle");
                    const payTotalEl = document.getElementById("paymentTotal");
                    if (payPlanEl) payPlanEl.textContent = plan;
                    if (payCycleEl) payCycleEl.textContent = cycle;
                    if (payTotalEl) payTotalEl.textContent = `₹${price}`;

                    AppState.pendingPlan = plan;
                    AppState.pendingPrice = price;
                    AppState.pendingCycle = cycle;

                    ViewRouter.show("Payment");
                });
            });

            // Back button
            document.getElementById("btnPlansBack")?.addEventListener("click", () => {
                const currentPlan = (AppState.subscriptionPlan || "").toLowerCase();
                const actualEndTs = AppState.trialEndsAt;
                const isExpired = currentPlan === 'expired' || (currentPlan.includes('trial') && actualEndTs && Date.now() >= actualEndTs);

                if (isExpired) {
                    const modal = document.getElementById("trialExpiredModal");
                    if (modal && modal.style.display !== "flex") {
                        modal.style.display = "flex";
                    }
                    ViewRouter.show("Dashboard");
                } else if (AppState.hasSubscription) {
                    ViewRouter.show("Dashboard");
                } else {
                    ViewRouter.show("Welcome");
                }
            });
        },

        // ---- Payment View ----
        bindPaymentView() {
            document.getElementById("btnPaymentBack")?.addEventListener("click", () => {
                ViewRouter.show("Plans");
            });

            document.getElementById("btnOpenCheckout")?.addEventListener("click", () => {
                CheckoutService.openCheckout(
                    AppState.pendingPlan,
                    AppState.pendingPrice,
                    AppState.pendingCycle
                );
            });

            document.getElementById("btnVerifyPayment")?.addEventListener("click", (e) => {
                e.preventDefault();
                CheckoutService.verifyPayment();
            });
        },

        // ---- Success View ----
        bindSuccessView() {
            document.getElementById("btnGotoDashboard")?.addEventListener("click", () => {
                DashboardService.render();
                ViewRouter.show("Dashboard");
            });
        },
    };
}
