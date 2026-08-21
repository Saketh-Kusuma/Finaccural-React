/**
 * Centralized reactive application state.
 *
 * Every value that has to survive a task-pane reload is mirrored into
 * localStorage under an `fa_` key and re-read here at module load, so the
 * object is always seeded from the last known session before any view
 * renders. Nothing in this module talks to the network or the DOM.
 */

const AppState = {
    // Auth / Subscription
    userEmail: localStorage.getItem("fa_user_email") || null,
    userName: localStorage.getItem("fa_user_name") || null,
    userProvider: localStorage.getItem("fa_user_provider") || null,
    hasSubscription: localStorage.getItem("fa_has_subscription") === "true",
    subscriptionId: localStorage.getItem("fa_subscription_id") || null,
    subscriptionPlan: (v => (!v || v === 'null' || v === 'undefined') ? null : v)(localStorage.getItem("fa_subscription_plan")),
    // Server-issued trial expiry (ms epoch), set at signup as
    // trial_ends_at and fetched via /api/auth/me. This is the
    // authoritative clock for the trial-expired popup — see
    // AppController.checkTrialExpiration().
    trialEndsAt: (v => (v ? parseInt(v, 10) : null))(localStorage.getItem("fa_trial_ends_at")),

    // Pending checkout details (set when user selects a plan)
    pendingPlan: null,
    pendingPrice: null,
    pendingCycle: null,

    // ERP Connection
    erpConnected: localStorage.getItem("fa_erp_connected") === "true",
    erpType: localStorage.getItem("fa_erp_type") || null,        // "quickbooks" | "xero"
    erpOrgName: localStorage.getItem("fa_erp_org") || null,
    erpConnectedDate: localStorage.getItem("fa_erp_date") || null,
    forceWelcome: false,

    // JWT token — persisted across sessions
    jwtToken: localStorage.getItem("fa_jwt_token") || null,
    // Opaque refresh token — used to renew the short-lived access JWT
    refreshToken: localStorage.getItem("fa_refresh_token") || null,
    // Guard flag: true while an access-token refresh is in progress so
    // concurrent 401s don't trigger multiple simultaneous refresh calls.
    _refreshing: false,
    _refreshQueue: [],
    // While true, ApiService.apiFetch refuses to make further requests
    // (avoids hammering the backend with a storm of repeated 401s)
    // until a fresh token is obtained via login.
    sessionExpired: false,

    // ERP Operations
    currentProvider: "quickbooks",
    get currentTier() {
        const plan = (AppState.subscriptionPlan || "pro").toLowerCase();
        if (plan.includes("trial")) return "trial";
        if (plan.includes("basic")) return "basic";
        if (plan.includes("standard")) return "standard";
        return "pro";
    },
    connectionId: null,
    isConnected: false,

    // Guards the "Redirecting to <provider>..." transition card
    // (launchERPOAuth) against repeat clicks while it's on screen —
    // true from the moment Connect is clicked until the popup either
    // opens, fails to open, or the user cancels via the close button.
    erpAuthInProgress: false
};

// Maximum connected companies (per ERP platform) allowed for a given
// plan name. Matches PLAN_LIMITS in Backend/src/modules/{quickbooks,xero}/service.js
// — keep both in sync. Uses `.includes()` so it matches both the short
// backend tier value ("trial", "basic", "standard") and any longer
// display label (e.g. "Free Trial (2 Hours)").
function getMaxCompaniesForPlan(plan) {
    const p = (plan || "").toLowerCase();
    if (p.includes("trial")) return 1;
    if (p.includes("basic")) return 1;
    if (p.includes("standard")) return 3;
    return 10;
}

export { AppState, getMaxCompaniesForPlan };
