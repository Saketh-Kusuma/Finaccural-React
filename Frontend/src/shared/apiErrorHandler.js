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
    const banner = document.getElementById("faGlobalBanner");
    if (banner) banner.style.display = "none";
}
