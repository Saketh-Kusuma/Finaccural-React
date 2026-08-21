/**
 * Checkout / payment popup flow and post-payment verification.
 */
import { AppState } from "../state/appState.js";
import { ViewRouter } from "../router/viewRouter.js";
import { ApiService } from "./apiService.js";
import { AuthService } from "./authService.js";
import { DashboardService } from "./dashboardService.js";

const CheckoutService = {
    /**
     * Initiates the mock hosted checkout flow.
     * In production: replace mock URL with Stripe/Razorpay checkout session URL.
     */
    openCheckout(plan, price, cycle) {
        AppState.pendingPlan = plan;
        AppState.pendingPrice = price;
        AppState.pendingCycle = cycle;

        const tokenParam = AppState.jwtToken ? `&token=${encodeURIComponent(AppState.jwtToken)}` : "";
        const checkoutUrl = `${ApiService.BASE}/api/payments/checkout?plan=${encodeURIComponent(plan)}&price=${price}&cycle=${encodeURIComponent(cycle)}&email=${encodeURIComponent(AppState.userEmail || "")}${tokenParam}`;

        const btnText = document.getElementById("checkoutBtnText");
        const btnSpinner = document.getElementById("checkoutSpinner");
        if (btnText) btnText.textContent = "Opening Secure Checkout...";
        if (btnSpinner) btnSpinner.classList.remove("hidden");

        const msgHandler = (event) => {
            if (!event.data) return;
            let data = event.data;
            if (typeof data === "string") {
                try { data = JSON.parse(data); } catch (_) { }
            }
            if (data && (data.type === "payment_success" || data.type === "checkout_complete")) {
                window.removeEventListener("message", msgHandler);
                CheckoutService.handlePaymentSuccess(data);
            }
        };
        window.addEventListener("message", msgHandler);

        const popup = window.open(
            checkoutUrl, "fa_checkout",
            "width=540,height=700,top=60,left=80,toolbar=no,menubar=no"
        );

        if (btnText) btnText.textContent = "Open Secure Checkout";
        if (btnSpinner) btnSpinner.classList.add("hidden");

        if (!popup || popup.closed) {
            window.removeEventListener("message", msgHandler);
            DashboardService.showError("Checkout popup was blocked. Please allow popups.");
        } else {
            popup.focus();
        }
    },

    /**
     * Handles successful payment message from checkout popup.
     * In production, backend verifies and returns subscription details.
     */
    handlePaymentSuccess(data) {
        const subId = data.subscriptionId || ("FA-SUB-" + Math.floor(100000 + Math.random() * 900000));
        const plan = data.plan || AppState.pendingPlan || "Professional";

        AppState.hasSubscription = true;
        AppState.subscriptionId = subId;
        AppState.subscriptionPlan = plan;
        AuthService._persistSubscription();

        // Show success screen
        const idEl = document.getElementById("successSubId");
        const planEl = document.getElementById("successPlanName");
        if (idEl) idEl.textContent = subId;
        if (planEl) planEl.textContent = plan;

        ViewRouter.show("Success");
        DashboardService.showStatus("Payment successful.", "success", `Subscribed to the ${plan} plan.`);
    },

    /**
     * Manually verifies a payment when user clicks "Verify my payment".
     * In production: calls GET /api/payments/verify?email=...
     */
    async verifyPayment() {
        ViewRouter.show("Loading");
        try {
            const res = await ApiService.checkSubscription(AppState.userEmail);
            if (res.hasSubscription) {
                AppState.hasSubscription = true;
                AppState.subscriptionId = res.subscriptionId || AppState.subscriptionId;
                AppState.subscriptionPlan = res.plan || AppState.pendingPlan;
                AuthService._persistSubscription();

                const idEl = document.getElementById("successSubId");
                const planEl = document.getElementById("successPlanName");
                if (idEl) idEl.textContent = AppState.subscriptionId;
                if (planEl) planEl.textContent = AppState.subscriptionPlan;
                ViewRouter.show("Success");
                DashboardService.showStatus("Payment successful.", "success", `Subscribed to the ${AppState.subscriptionPlan} plan.`);
            } else {
                ViewRouter.show("Payment");
                DashboardService.showStatus("Payment verification failed.", "error", "We couldn't find an active subscription yet. Please try again.");
            }
        } catch {
            ViewRouter.show("Payment");
            DashboardService.showStatus("Payment verification failed.", "error", "Please try again.");
        }
    }
};

export { CheckoutService };
