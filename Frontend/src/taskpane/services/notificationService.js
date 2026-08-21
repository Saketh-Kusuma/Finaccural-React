/**
 * Bell icon / drawer notification history.
 *
 * Entries are owned by the backend (`/api/notifications`) and mirrored in a
 * local cache so the drawer can render synchronously; everything shown is
 * scoped to the active ERP provider + company via `_forCurrentContext()`.
 */
import { AppState } from "../state/appState.js";
import { ApiService } from "./apiService.js";

const NotificationService = {
    MAX_ITEMS: 50,
    _lastNotif: null,

    // In-memory cache of the logged-in user's notification history,
    // as last fetched from the backend (GET /api/notifications). The
    // backend is now the source of truth — this cache just avoids an
    // API round-trip on every renderBadge()/renderDrawer() call.
    // Populated by init() on load and kept in sync by every
    // add()/markAllRead()/clearAll() mutation.
    _cache: [],

    /**
     * Normalizes a provider tag to "quickbooks" | "xero" | null.
     * null means "global" — not tied to either ERP, so it's always
     * visible regardless of which provider the user is currently on
     * (e.g. login, payment, logout).
     * @param {string} [provider]
     * @returns {"quickbooks"|"xero"|null}
     */
    _normalizeProvider(provider) {
        if (provider === "quickbooks" || provider === "xero") return provider;
        return null;
    },

    /**
     * Filters a notification list down to what's visible in the
     * current provider context: global (provider-less) entries plus
     * whichever ERP the user is actively using. This is the single
     * choke point that keeps QuickBooks and Xero notifications
     * (toast, badge count, and drawer history) fully separated —
     * QuickBooks notifications never surface while on Xero and vice
     * versa.
     * @param {Array} list
     * @returns {Array}
     */
    _forCurrentContext(list) {
        const ctx = (typeof AppState !== "undefined" && AppState.currentProvider) || null;
        return list.filter(n => !n.provider || n.provider === ctx);
    },

    /**
     * Maps a backend notification row (modules/notifications/model —
     * id, userId, type, message, detail, provider, read, createdAt) to
     * the shape the rest of this service already expects.
     * @returns {{id:string,type:'success'|'error',message:string,detail:string,provider:('quickbooks'|'xero'|null),timestamp:string,read:boolean}}
     */
    _mapFromBackend(row) {
        return {
            id: row.id,
            type: row.type === "error" ? "error" : "success",
            message: row.message || "",
            detail: row.detail || "",
            provider: this._normalizeProvider(row.provider),
            timestamp: row.createdAt || row.created_at || new Date().toISOString(),
            read: !!row.read
        };
    },

    /**
     * Fetches the logged-in user's full notification history from the
     * backend (GET /api/notifications, newest first) and refreshes the
     * in-memory cache. Failures (offline, logged out, etc.) are
     * swallowed — ApiService.apiFetch already surfaces the global
     * offline/session banners, and the bell/drawer just keep showing
     * whatever was last cached rather than throwing.
     * @returns {Promise<Array>}
     */
    async _fetchFromBackend() {
        try {
            const res = await ApiService.apiFetch("/api/notifications");
            if (!res.ok) return this._cache;
            const data = await res.json();
            const rows = Array.isArray(data.notifications) ? data.notifications : [];
            this._cache = rows.map(r => this._mapFromBackend(r));
        } catch (_) {
            // Network/parse failure — keep whatever was cached before.
        }
        return this._cache;
    },

    /**
     * Creates a notification server-side.
     * @returns {Promise<object|null>} the created row, or null on failure
     */
    async _postToBackend(type, message, detail, provider) {
        try {
            const res = await ApiService.apiFetch("/api/notifications", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    type,
                    message,
                    detail: detail ? String(detail) : undefined,
                    provider: provider || undefined
                })
            });
            if (!res.ok) return null;
            const data = await res.json();
            return data.notification || null;
        } catch (_) {
            return null;
        }
    },

    _escapeHtml(str) {
        const div = document.createElement("div");
        div.textContent = String(str == null ? "" : str);
        return div.innerHTML;
    },

    _formatTimestamp(iso) {
        const d = new Date(iso);
        if (isNaN(d.getTime())) return "";
        const now = new Date();
        const timeStr = d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
        if (d.toDateString() === now.toDateString()) return timeStr;
        const dateStr = d.toLocaleDateString([], { month: "short", day: "numeric" });
        return `${dateStr}, ${timeStr}`;
    },

    /**
     * Records a new notification, shows it immediately as a top-right
     * toast (the primary way the user learns an action succeeded or
     * failed — the bell is just the persisted history, never required
     * reading), and refreshes the bell badge. If the drawer happens to
     * already be open, the new entry is shown there too and marked read.
     *
     * QuickBooks/Xero separation: `provider` tags which ERP the action
     * belongs to. A tagged notification only ever toasts, counts toward
     * the badge, or appears in the drawer while the user is actively on
     * that same provider (AppState.currentProvider) — it's still saved
     * to history so it surfaces correctly if the user switches back
     * later, but it's fully invisible in the meantime. Untagged (global)
     * notifications — login, payment, logout, etc. — aren't tied to
     * either ERP and always show.
     * @param {string} message - short outcome, e.g. "Data completed."
     * @param {"success"|"error"} type
     * @param {string} [detail] - optional second line, e.g. "Data synchronized successfully."
     * @param {"quickbooks"|"xero"} [provider] - which ERP this belongs to, if any
     */
    add(message, type, detail, provider) {
        if (!message) return;
        const normalizedType = type === "error" ? "error" : "success";
        const normalizedProvider = this._normalizeProvider(provider);

        // Duplicate guard — the same completed outcome can occasionally
        // be reported twice (e.g. two callback paths both observing the
        // same finished operation). Suppress an identical repeat within
        // a short window so each completed action produces exactly one
        // toast/notification, never more.
        const now = Date.now();
        if (
            this._lastNotif &&
            this._lastNotif.message === message &&
            this._lastNotif.type === normalizedType &&
            this._lastNotif.provider === normalizedProvider &&
            (now - this._lastNotif.at) < 1500
        ) {
            return;
        }
        this._lastNotif = { message, type: normalizedType, provider: normalizedProvider, at: now };

        // Only surface it (toast + badge/drawer refresh) if it's global
        // or matches the ERP the user is currently on — a QuickBooks
        // notification must never appear while on Xero, and vice versa.
        const currentCtx = (typeof AppState !== "undefined" && AppState.currentProvider) || null;
        const isVisibleNow = !normalizedProvider || normalizedProvider === currentCtx;

        // Immediate feedback — always, regardless of whether the drawer
        // is open or the bell has ever been clicked, and regardless of
        // the backend round-trip below (toast display is local-only and
        // must never wait on the network).
        if (isVisibleNow) {
            this.showToast(message, normalizedType, detail);
        }

        // Persist to the backend. add() itself stays synchronous/
        // non-blocking for callers (same contract every existing call
        // site already relies on) — the cache + badge/drawer refresh
        // just land a beat after the toast once the POST resolves.
        this._postToBackend(normalizedType, message, detail, normalizedProvider).then(row => {
            const entry = row ? this._mapFromBackend(row) : {
                id: `local_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
                type: normalizedType,
                message: String(message),
                detail: detail ? String(detail) : "",
                provider: normalizedProvider,
                timestamp: new Date().toISOString(),
                read: false
            };
            this._cache.unshift(entry);
            if (this._cache.length > this.MAX_ITEMS) this._cache.length = this.MAX_ITEMS;

            if (!isVisibleNow) return;

            const drawer = document.getElementById("notifDrawer");
            if (drawer && drawer.style.display !== "none") {
                // Drawer is open — the new notification is on screen, so
                // treat it as read immediately instead of leaving a stray
                // unread badge behind an already-open drawer.
                this.markAllRead();
                this.renderDrawer();
            } else {
                this.renderBadge();
            }
        });
    },

    /**
     * Renders a single floating, top-right toast card. Multiple toasts
     * stack (most recent on top); each auto-dismisses on its own timer
     * but can also be closed early via the × button.
     * @param {string} message
     * @param {"success"|"error"} type
     * @param {string} [detail]
     * @param {number} [duration=4500]
     */
    showToast(message, type, detail, duration = 4500) {
        const container = document.getElementById("notifToastContainer");
        if (!container) return;

        const toastEl = document.createElement("div");
        toastEl.className = `fa-notif-toast fa-notif-toast-${type === "error" ? "error" : "success"}`;
        toastEl.innerHTML = `
                <span class="fa-notif-toast-icon">${type === "error" ? "✕" : "✓"}</span>
                <div class="fa-notif-toast-body">
                    <div class="fa-notif-toast-title">${this._escapeHtml(message)}</div>
                    ${detail ? `<div class="fa-notif-toast-detail">${this._escapeHtml(detail)}</div>` : ""}
                </div>
                <button class="fa-notif-toast-close" aria-label="Close">&times;</button>
            `;
        container.appendChild(toastEl);

        let dismissTimer = null;
        const removeToast = () => {
            clearTimeout(dismissTimer);
            toastEl.classList.add("fa-notif-toast-hide");
            setTimeout(() => toastEl.remove(), 200);
        };

        toastEl.querySelector(".fa-notif-toast-close")?.addEventListener("click", removeToast);
        dismissTimer = setTimeout(removeToast, duration);
    },

    /**
     * Re-renders the badge (and the drawer, if it's open) against the
     * current provider context. Call this whenever AppState.currentProvider
     * changes (switching between QuickBooks and Xero) so a badge count
     * or open drawer left over from the other provider doesn't linger —
     * it recomputes to show only what's actually visible now.
     */
    refreshForContext() {
        this.renderBadge();
        const drawer = document.getElementById("notifDrawer");
        if (drawer && drawer.style.display !== "none") {
            this.renderDrawer();
        }
    },

    /** Notifications visible right now — global entries plus the active provider's. Reads the in-memory cache (see _fetchFromBackend). */
    getAll() {
        return this._forCurrentContext(this._cache);
    },

    getUnreadCount() {
        return this._forCurrentContext(this._cache).filter(n => !n.read).length;
    },

    /**
     * Marks read only the notifications currently visible (global +
     * active provider) — a QuickBooks notification the user hasn't
     * seen yet (because they're on Xero) stays unread until they
     * actually switch to QuickBooks and see it. Updates the local
     * cache immediately (so the badge clears without waiting on the
     * network) and fires the PATCH in the background.
     */
    markAllRead() {
        const ctx = (typeof AppState !== "undefined" && AppState.currentProvider) || null;
        const ids = [];
        this._cache.forEach(n => {
            if (!n.read && (!n.provider || n.provider === ctx)) {
                n.read = true;
                ids.push(n.id);
            }
        });
        this.renderBadge();

        if (!ids.length) return;
        ApiService.apiFetch("/api/notifications/mark-read", {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ ids })
        }).catch(() => {
            // Best-effort — a later _fetchFromBackend() resync (e.g. the
            // next bell open after a taskpane reload) corrects any drift.
        });
    },

    /**
     * Permanently deletes every notification (all providers) for the
     * logged-in user — backend first, then cache.
     *
     * Deliberately NOT optimistic: apiFetch() resolves normally (not a
     * rejected promise) for a non-2xx response — a bare `.catch()`
     * around the DELETE call would never fire on e.g. a 401/500, so a
     * failed delete would silently look successful client-side while
     * the rows stayed in the DB and reappeared for this user on the
     * next _fetchFromBackend() (next login/reload). Checking res.ok
     * here and only clearing the cache once the backend confirms the
     * delete is what actually fixes that "Clear All doesn't clear"
     * case, instead of just hiding it for the rest of this session.
     */
    async clearAll() {
        try {
            const res = await ApiService.apiFetch("/api/notifications", { method: "DELETE" });
            if (!res.ok) {
                this.showToast("Couldn't clear notifications. Please try again.", "error");
                return;
            }
        } catch (_) {
            this.showToast("Couldn't clear notifications. Please try again.", "error");
            return;
        }

        this._cache = [];
        this.renderBadge();
        this.renderDrawer();
    },

    renderBadge() {
        const count = this.getUnreadCount();
        const text = count > 99 ? "99+" : String(count);
        const displayStyle = count > 0 ? "flex" : "none";
        
        const badges = document.querySelectorAll(".fa-notif-badge");
        badges.forEach(badge => {
            badge.textContent = text;
            badge.style.display = displayStyle;
        });
    },

    renderDrawer() {
        const listEl = document.getElementById("notifList");
        if (!listEl) return;
        const notifications = this.getAll();

        if (notifications.length === 0) {
            listEl.innerHTML = `<div class="fa-notif-empty">No notifications available.</div>`;
            return;
        }

        listEl.innerHTML = notifications.map(n => {
            const icon = n.type === "error" ? "✕" : "✓";
            return `
                    <div class="fa-notif-item fa-notif-item-${n.type === "error" ? "error" : "success"}">
                        <span class="fa-notif-icon">${icon}</span>
                        <div class="fa-notif-content">
                            <div class="fa-notif-message">${this._escapeHtml(n.message)}</div>
                            ${n.detail ? `<div class="fa-notif-detail">${this._escapeHtml(n.detail)}</div>` : ""}
                            <div class="fa-notif-time">${this._formatTimestamp(n.timestamp)}</div>
                        </div>
                    </div>
                `;
        }).join("");
    },

    /** Wires up the bell button, drawer, and Clear All, then loads notification history from the backend. Call once on init. */
    async init() {
        this.renderBadge();

        // The notification history now lives server-side, scoped per
        // user (see modules/notifications). Drop any leftover
        // `fa_notifications` data from before this migration so it can
        // never leak between different users signed into the same
        // browser — it's simply ignored/cleared on first login.
        try {
            localStorage.removeItem("fa_notifications");
        } catch (_) { /* ignore */ }

        const drawer = () => document.getElementById("notifDrawer");

        const toggleDrawer = (e) => {
            e.stopPropagation();
            // Close account menu dropdown first to prevent overlap
            const dropdown = document.getElementById("accountMenuDropdown");
            if (dropdown) dropdown.style.display = "none";

            const el = drawer();
            if (!el) return;
            const willOpen = el.style.display === "none";
            if (willOpen) {
                // Set correct top offset based on whether connected dashboard has row 2 visible
                const isConnected = document.getElementById("dashConnected")?.style.display !== "none";
                el.style.top = isConnected ? "88px" : "56px";
                
                // Opening the drawer shows the full history and marks
                // everything read — the unread badge disappears.
                this.renderDrawer();
                this.markAllRead();
            }
            el.style.display = willOpen ? "flex" : "none";
        };
        document.getElementById("notifBellBtn1")?.addEventListener("click", toggleDrawer);
        document.getElementById("notifBellBtn2")?.addEventListener("click", toggleDrawer);
        document.getElementById("notifBellBtn3")?.addEventListener("click", toggleDrawer);

        // Close the drawer when clicking anywhere outside it (or any of the bells).
        document.addEventListener("click", (e) => {
            const el = drawer();
            if (el && el.style.display !== "none" && !el.contains(e.target)) {
                // Check if the click was on any of the bells
                const clickedBell = e.target.closest(".fa-notif-bell-btn");
                if (!clickedBell) {
                    el.style.display = "none";
                }
            }
        });

        // Clear All — deletes immediately, no confirmation dialog, then
        // closes the drawer. Stop propagation so this click doesn't
        // also trigger the "click outside closes the drawer" listener
        // above (that's a no-op here anyway since we close it
        // ourselves, but keeps behavior explicit/predictable).
        document.getElementById("notifClearAllBtn")?.addEventListener("click", (e) => {
            e.stopPropagation();
            this.clearAll();
            const el = drawer();
            if (el) el.style.display = "none";
        });

        // Load this user's notification history from the backend and
        // render the real badge count. Buttons above are already wired
        // and usable while this is in flight.
        await this._fetchFromBackend();
        this.renderBadge();
    }
};

export { NotificationService };
