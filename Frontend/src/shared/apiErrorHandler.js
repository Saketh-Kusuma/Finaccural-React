/**
 * Centralized API error handling for the taskpane.
 * -----------------------------------------------------------------
 * Pairs with the backend's centralized Express error middleware
 * (Backend/Backend/src/core/middleware/errorHandler.js). Every
 * non-2xx response from the backend follows the same envelope:
 *   { success: false, code, message, details }
 *
 * This module owns:
 *   - ApiError            : a typed error carrying `.code`, so callers
 *                            branch on `code`, never on message text.
 *   - parseApiError()      : safely turns a failed fetch Response into
 *                            an ApiError, tolerating non-JSON bodies.
 *   - networkError()        : turns a raw fetch() network failure
 *                            (backend unreachable / offline) into the
 *                            same ApiError shape.
 *   - showBanner/hideBanner : the persistent, actionable banner used for
 *                            the offline (red) and ERP-expired (orange)
 *                            scenarios.
 *   - showToast            : the transient toast used for session-expired.
 *
 * It deliberately does NOT know about AppState/ViewRouter/AuthService —
 * taskpane.js decides *what to do* (redirect, retry, reconnect); this
 * module only knows *how to show it*, which keeps it reusable for any
 * future API call without creating a circular dependency.
 * -----------------------------------------------------------------
 */

import { ERROR_CODES, getFriendlyMessage } from "./errorMessages.js";

export { ERROR_CODES };

export class ApiError extends Error {
    constructor(code, message, details) {
        super(message);
        this.name = "ApiError";
        this.code = code || ERROR_CODES.UNKNOWN;
        // `details` is technical/log-only — callers must never render this
        // in the UI. Log it to the console for developers, nothing more.
        this.details = details || message;
    }
}

/**
 * Parses a non-ok fetch Response using the backend's standardized error
 * envelope. Falls back gracefully if the body isn't JSON (e.g. a proxy's
 * HTML error page) or doesn't include the expected fields, so callers
 * never need their own try/catch around this.
 * @param {Response} response
 * @returns {Promise<ApiError>}
 */
export async function parseApiError(response) {
    let body = {};
    try {
        body = await response.clone().json();
    } catch (_) {
        // Non-JSON body — leave body as {} and fall back below.
    }

    const code = body.code || (response.status === 401 ? ERROR_CODES.SESSION_EXPIRED : ERROR_CODES.UNKNOWN);
    const message = body.message || getFriendlyMessage(code);
    const details = body.details || `HTTP ${response.status} ${response.statusText}`;

    const err = new ApiError(code, message, details);
    err.status = response.status;
    // Technical detail goes to the console only — never the UI.
    console.error(`[API ${response.status}] ${code}:`, details);
    return err;
}

/**
 * Turns a raw fetch() failure (server unreachable, DNS failure, no
 * internet, connection refused — i.e. no HTTP response at all) into the
 * same standardized ApiError shape as a backend error response.
 * @param {Error} originalError
 * @returns {ApiError}
 */
export function networkError(originalError) {
    const message = getFriendlyMessage(ERROR_CODES.CONNECTION_REFUSED);
    console.error("[API] Network error:", (originalError && originalError.message) || originalError);
    return new ApiError(
        ERROR_CODES.CONNECTION_REFUSED,
        message,
        (originalError && originalError.message) || "Network request failed."
    );
}

// ---------------------------------------------------------------
// UI primitives — persistent banner (offline / ERP expired) + toast
// (session expired). Pure DOM helpers; no app-state knowledge.
// ---------------------------------------------------------------

let toastTimer = null;

/** Transient toast notification, auto-dismisses after `duration` ms. */
export function showToast(message, duration = 4000) {
    if (typeof window !== "undefined" && window._faIsOffline) return;
    const banner = document.getElementById("faGlobalBanner");
    if (banner && banner.style.display !== "none" && banner.style.display !== "") return;

    const el = document.getElementById("faToast");
    if (!el) return;
    el.textContent = message;
    el.style.display = "block";
    // Force reflow so the transition re-triggers on repeated calls.
    void el.offsetWidth;
    el.classList.add("fa-toast-visible");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
        el.classList.remove("fa-toast-visible");
        setTimeout(() => {
            if (!el.classList.contains("fa-toast-visible")) el.style.display = "none";
        }, 250);
    }, duration);
}

/**
 * Persistent, actionable banner pinned to the top of the taskpane.
 * @param {object} opts
 * @param {"offline"|"erp"|"error"} opts.type
 * @param {string} opts.message
 * @param {string} [opts.actionLabel]
 * @param {() => void} [opts.onAction]
 */
export function showBanner({ type = "error", message, actionLabel, onAction }) {
    const banner    = document.getElementById("faGlobalBanner");
    const iconEl     = document.getElementById("faGlobalBannerIcon");
    const msgEl      = document.getElementById("faGlobalBannerMessage");
    const actionEl   = document.getElementById("faGlobalBannerAction");
    const dismissEl  = document.getElementById("faGlobalBannerDismiss");
    if (!banner) return;

    if (type === "offline") {
        if (typeof window !== "undefined") window._faIsOffline = true;
    }

    banner.className = `fa-global-banner fa-global-banner-${type}`;
    banner.style.display = "flex";
    if (iconEl) iconEl.textContent = type === "offline" ? "📡" : type === "erp" ? "🔌" : "⚠️";
    if (msgEl) msgEl.textContent = message;

    if (actionEl) {
        if (actionLabel && onAction) {
            actionEl.textContent = actionLabel;
            actionEl.style.display = "inline-flex";
            actionEl.onclick = onAction;
        } else {
            actionEl.style.display = "none";
            actionEl.onclick = null;
        }
    }

    if (dismissEl) dismissEl.onclick = () => hideBanner();
}

export function hideBanner() {
    if (typeof window !== "undefined") window._faIsOffline = false;
    const banner = document.getElementById("faGlobalBanner");
    if (banner) banner.style.display = "none";
}

/** Flashes/shakes the banner to visually alert user when a click is blocked offline */
export function flashGlobalBanner() {
    const banner = document.getElementById("faGlobalBanner");
    if (banner) {
        banner.classList.remove("fa-banner-flash");
        void banner.offsetWidth; // Force reflow
        banner.classList.add("fa-banner-flash");
    }
}

/**
 * Checks connectivity status:
 * 1. Checks navigator.onLine and pings an internet endpoint (https://www.gstatic.com/generate_204)
 *    If fails -> "Offline: Please check your internet connection."
 * 2. If Internet is ON, pings backend server (http://localhost:8000/api/health)
 *    If fails -> "Offline: Cannot connect to the server."
 * 3. If online & server reachable -> hides banner and returns true.
 * @returns {Promise<boolean>}
 */
export async function checkConnectionStatus() {
    // Stage 1: Check browser network indicator
    if (typeof navigator !== "undefined" && !navigator.onLine) {
        if (typeof window !== "undefined") window._faIsOffline = true;
        showBanner({
            type: "offline",
            message: "Offline: Please check your internet connection.",
            actionLabel: "Retry",
            onAction: () => {
                checkConnectionStatus();
            }
        });
        return false;
    }

    // Stage 2: Internet Ping to verify actual Wi-Fi/Internet access beyond local loopback
    let isInternetAlive = false;
    try {
        const inetController = typeof AbortController !== "undefined" ? new AbortController() : null;
        const inetTimeout = inetController ? setTimeout(() => inetController.abort(), 2000) : null;
        await fetch("https://www.gstatic.com/generate_204", {
            method: "GET",
            mode: "no-cors",
            cache: "no-store",
            signal: inetController ? inetController.signal : undefined
        });
        if (inetTimeout) clearTimeout(inetTimeout);
        isInternetAlive = true;
    } catch (_) {
        isInternetAlive = false;
    }

    if (!isInternetAlive) {
        if (typeof window !== "undefined") window._faIsOffline = true;
        showBanner({
            type: "offline",
            message: "Offline: Please check your internet connection.",
            actionLabel: "Retry",
            onAction: () => {
                checkConnectionStatus();
            }
        });
        return false;
    }

    // Stage 3: Internet is ON! Check local backend server health
    try {
        const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
        const timeoutId = controller ? setTimeout(() => controller.abort(), 3500) : null;
        const res = await fetch("http://localhost:8000/api/health", {
            method: "GET",
            signal: controller ? controller.signal : undefined
        });
        if (timeoutId) clearTimeout(timeoutId);
        if (res.ok) {
            hideBanner();
            return true;
        }
    } catch (_) {}

    // Backend server is down
    if (typeof window !== "undefined") window._faIsOffline = true;
    showBanner({
        type: "offline",
        message: "Offline: Cannot connect to the server.",
        actionLabel: "Retry",
        onAction: () => {
            checkConnectionStatus();
        }
    });
    return false;
}

// ---------------------------------------------------------------
// Global connectivity watchers & Capture-phase click interceptor
// ---------------------------------------------------------------
let isCheckingConnection = false;

if (typeof window !== "undefined") {
    window.addEventListener("online", () => {
        checkConnectionStatus();
    });
    window.addEventListener("offline", () => {
        window._faIsOffline = true;
        checkConnectionStatus();
    });

    document.addEventListener("click", (e) => {
        const banner = document.getElementById("faGlobalBanner");
        const targetBtn = e.target.closest("button, a, select, input[type='button'], input[type='submit'], [role='button'], .fa-company-row, .fa-modal-company-row, .erp-card-large, .conn-journal-btn, .fa-journal-btn");
        if (!targetBtn) return;

        // Allow clicks on the banner itself (e.g. Retry button or Dismiss X)
        if (banner && banner.contains(targetBtn)) return;

        const isBrowserOffline = typeof navigator !== "undefined" && !navigator.onLine;
        const isBannerShown = banner && banner.style.display !== "none" && banner.style.display !== "";
        const isOffline = isBrowserOffline || window._faIsOffline || isBannerShown;

        if (isOffline) {
            // Synchronously cancel event propagation to block all button handlers!
            e.preventDefault();
            e.stopPropagation();
            if (typeof e.stopImmediatePropagation === "function") {
                e.stopImmediatePropagation();
            }

            flashGlobalBanner();

            if (!isCheckingConnection) {
                isCheckingConnection = true;
                checkConnectionStatus().finally(() => {
                    isCheckingConnection = false;
                });
            }
            return false;
        }
    }, true); // Capture phase to intercept synchronously before component handlers!
}

