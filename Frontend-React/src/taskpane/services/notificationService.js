/**
 * Notification Service for FinAccrual React
 * Syncs with backend /api/notifications and manages notifications list & badges.
 */
import { apiFetch } from "../api";

export const NotificationService = {
  MAX_ITEMS: 50,

  _normalizeProvider(provider) {
    if (provider === "quickbooks" || provider === "xero") return provider;
    return null;
  },

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

  async fetchNotifications() {
    try {
      const res = await apiFetch("/api/notifications");
      if (!res.ok) return [];
      const data = await res.json();
      const rows = Array.isArray(data.notifications) ? data.notifications : [];
      return rows.map((r) => this._mapFromBackend(r));
    } catch (_) {
      return [];
    }
  },

  async postNotification(type, message, detail, provider) {
    try {
      const res = await apiFetch("/api/notifications", {
        method: "POST",
        body: JSON.stringify({
          type,
          message,
          detail: detail ? String(detail) : undefined,
          provider: provider || undefined
        })
      });
      if (!res.ok) return null;
      const data = await res.json();
      return data.notification ? this._mapFromBackend(data.notification) : null;
    } catch (_) {
      return null;
    }
  },

  async markRead(ids) {
    if (!ids || ids.length === 0) return;
    try {
      await apiFetch("/api/notifications/mark-read", {
        method: "PATCH",
        body: JSON.stringify({ ids })
      });
    } catch (_) {}
  },

  async clearAll() {
    try {
      const res = await apiFetch("/api/notifications", { method: "DELETE" });
      return res.ok;
    } catch (_) {
      return false;
    }
  }
};
