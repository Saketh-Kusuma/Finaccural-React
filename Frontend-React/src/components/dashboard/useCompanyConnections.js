import { useState, useEffect, useCallback } from "react";
import { apiFetch, openErp, isTrustedOrigin } from "../../taskpane/api";
import { ExcelService } from "../../taskpane/services/excelService";

export function formatRelativeTime(dateInput, status) {
  if (status === "Disconnected") return "Disconnected";
  if (status === "Not Synced" || !dateInput) return "Not Synced";
  const date = new Date(dateInput);
  if (isNaN(date.getTime())) return "Not Synced";
  const diffMs = new Date() - date;
  if (diffMs < 0) return "Just now";
  const diffSec = Math.floor(diffMs / 1000);
  if (diffSec < 45) return "Just now";
  const diffMin = Math.floor(diffSec / 60);
  if (diffMin < 60) return diffMin === 1 ? "1 minute ago" : `${diffMin} minutes ago`;
  const diffHr = Math.floor(diffMin / 60);
  if (diffHr < 24) return diffHr === 1 ? "1 hour ago" : `${diffHr} hours ago`;
  const diffDays = Math.floor(diffHr / 24);
  if (diffDays < 30) return diffDays === 1 ? "1 day ago" : `${diffDays} days ago`;
  return date.toLocaleDateString();
}

export function useCompanyConnections({
  provider,
  user,
  planClean,
  notify,
  addLog,
  onConnect,
  onChangePlan,
}) {
  const isXero = (provider || "").toLowerCase() === "xero";
  const label = isXero ? "Xero" : "QuickBooks";

  const maxCompanies = planClean.toLowerCase().includes("pro")
    ? 10
    : planClean.toLowerCase().includes("standard")
      ? 3
      : 1;

  const [connections, setConnections] = useState([]);
  const [activeCompanyId, setActiveCompanyId] = useState(
    () => localStorage.getItem("fa_current_company_id") || null
  );

  const reloadConnections = useCallback(async () => {
    const email = user.email || localStorage.getItem("fa_user_email") || "";
    if (!email) return [];
    try {
      const res = await apiFetch(`/api/connections?mail=${encodeURIComponent(email)}`);
      const data = await res.json();
      if (Array.isArray(data)) {
        setConnections(data);
        return data;
      }
    } catch (_) {}
    return [];
  }, [user.email]);

  useEffect(() => {
    let mounted = true;
    const email = user.email || localStorage.getItem("fa_user_email") || "";
    if (email) {
      apiFetch(`/api/connections?mail=${encodeURIComponent(email)}`)
        .then((res) => res.json())
        .then((data) => {
          if (mounted && Array.isArray(data) && data.length > 0) {
            setConnections(data);
          }
        })
        .catch(() => {});
    }
    return () => {
      mounted = false;
    };
  }, [user.email]);

  useEffect(() => {
    const receive = (event) => {
      if (!isTrustedOrigin(event.origin)) return;
      if (event.data === "qb_connected" || event.data === "xero_connected") {
        if (addLog) addLog(`Connection completed: ${event.data}`);
        const pendingReconnectId = typeof sessionStorage !== "undefined"
          ? sessionStorage.getItem("fa_pending_reconnect_id")
          : null;
        if (typeof sessionStorage !== "undefined") {
          sessionStorage.removeItem("fa_pending_reconnect_id");
        }

        reloadConnections().then((conns) => {
          if (pendingReconnectId) {
            const reconnected = conns.find((c) => c.companyId === pendingReconnectId);
            if (reconnected) {
              setActiveCompanyId(reconnected.companyId);
              localStorage.setItem("fa_current_company_id", reconnected.companyId);
              notify("Company reconnected successfully.", "success", `${reconnected.companyName || label} is now active and re-authorized.`, provider);
              return;
            }
          }
          const matching = conns.filter(
            (c) => (c.platform || "").toLowerCase() === (provider || "").toLowerCase()
          );
          if (matching.length > 0) {
            const newest = matching[matching.length - 1];
            setActiveCompanyId(newest.companyId);
            localStorage.setItem("fa_current_company_id", newest.companyId);
          }
        });
      }
    };
    window.addEventListener("message", receive);
    return () => window.removeEventListener("message", receive);
  }, [provider, reloadConnections, addLog, label, notify]);

  useEffect(() => {
    const handleErpExpired = (event) => {
      const details = event?.detail || {};
      if (addLog) addLog(`ERP session expired: ${details.message || "Please reconnect company."}`);
      notify("Connection Expired", "error", details.message || `Your ${label} connection has expired. Please click Reconnect to restore access.`, provider);
      reloadConnections();
    };
    window.addEventListener("fa_erp_session_expired", handleErpExpired);
    return () => window.removeEventListener("fa_erp_session_expired", handleErpExpired);
  }, [label, provider, addLog, notify, reloadConnections]);

  const platformConns = connections.filter(
    (c) => (c.platform || "").toLowerCase() === (provider || "").toLowerCase()
  );

  const activeConnection =
    platformConns.find((c) => c.companyId === activeCompanyId) ||
    platformConns[0] ||
    connections.find(
      (c) => (c.platform || "").toLowerCase() === (provider || "").toLowerCase()
    ) ||
    connections[0];

  const companyName =
    activeConnection?.companyName ||
    (isXero ? "sushanth" : "Sandbox Company GB ce1f");
  const realmId =
    activeConnection?.companyId || "c90dc421-681f-4bc6-a3c0-812a6c44ab03";
  const lastSyncText = formatRelativeTime(
    activeConnection?.lastSyncedAt,
    activeConnection?.status
  );

  const connectedCount =
    platformConns.length > 0
      ? platformConns.length
      : connections.length > 0
        ? connections.length
        : 1;
  const remainingCompanies = Math.max(0, maxCompanies - connectedCount);

  const handleAddCompanyClick = async () => {
    try {
      await ExcelService.clearMasterData();
    } catch (err) {
      console.error("Error clearing Excel data: ", err);
    }

    const email = user.email || localStorage.getItem("fa_user_email") || "";
    let conns = connections;
    try {
      const res = await apiFetch("/api/connections?mail=" + encodeURIComponent(email));
      const data = await res.json();
      if (Array.isArray(data)) {
        conns = data;
        setConnections(data);
      }
    } catch (_) {}

    const connsForProvider = (conns || []).filter(
      (c) => (c.platform || "").toLowerCase() === (provider || "quickbooks").toLowerCase()
    );

    if (connsForProvider.length >= maxCompanies) {
      const alertMsg = `Your ${planClean} Plan limits active ${label} connections to ${maxCompanies} companies. Please upgrade your plan to connect more companies.`;
      notify(`Company limit reached (${connsForProvider.length}/${maxCompanies})`, "error", alertMsg, provider);
      if (onChangePlan) {
        onChangePlan();
      }
      return;
    }

    if (addLog) addLog(`Opening ${label} connection window...`);
    notify(`Connecting to ${label}...`, "success", null, provider);

    if (onConnect) {
      onConnect(provider);
    } else {
      const popup = openErp(provider, user);
      if (!popup) {
        notify("Connection window blocked. Please allow popups.", "error", null, provider);
      }
    }
  };

  const switchActiveCompany = async (companyId) => {
    setActiveCompanyId(companyId);
    localStorage.setItem("fa_current_company_id", companyId);
    try {
      await ExcelService.clearMasterData();
    } catch (err) {
      console.error("Error clearing Excel data: ", err);
    }

    const targetConn = platformConns.find((c) => c.companyId === companyId);
    try {
      const res = await apiFetch(`/api/connections/${companyId}/activate`, { method: "POST" });
      const data = await res.json().catch(() => ({}));
      if (targetConn) {
        const countMsg = data.totalRecords ? ` (Total Records Found: ${data.totalRecords})` : "";
        if (addLog) addLog(`Switched active company to: ${targetConn.companyName}${countMsg}`);
        notify(
          `Active company updated to ${targetConn.companyName}`,
          "success",
          data.totalRecords !== undefined ? `Total Records Found: ${data.totalRecords}` : null,
          provider
        );
      }
      reloadConnections();
    } catch (err) {
      console.error("Error activating company:", err);
      notify("Failed to switch active company.", "error", null, provider);
    }
  };

  const handleDisconnectCompany = async (company) => {
    const compId = company.companyId;
    const compName = company.companyName || "Company";
    if (addLog) addLog(`Disconnecting ${compName}...`);
    notify(`Disconnecting ${compName}...`, "success", null, provider);
    try {
      await apiFetch(`/api/connections/${compId}`, { method: "DELETE" });
      if (activeCompanyId === compId) {
        setActiveCompanyId(null);
        localStorage.removeItem("fa_current_company_id");
      }
      try {
        await ExcelService.clearMasterData();
      } catch (_) {}
      notify("Company disconnected.", "success", null, provider);
      if (addLog) addLog(`Company disconnected: ${compName}`);
      reloadConnections();
    } catch (_) {
      notify("Failed to disconnect company.", "error", null, provider);
    }
  };

  const handleRenameCompany = async (renamingCompany, newName) => {
    const trimmed = (newName || "").trim();
    if (!trimmed || !renamingCompany) return;
    try {
      await apiFetch(`/api/connections/${renamingCompany.companyId}/rename`, {
        method: "PATCH",
        body: JSON.stringify({ companyName: trimmed }),
      });
      notify("Company renamed successfully.", "success", null, provider);
      reloadConnections();
    } catch (_) {
      notify("Failed to rename company.", "error", null, provider);
    }
  };

  return {
    label,
    isXero,
    maxCompanies,
    connections,
    platformConns,
    activeCompanyId,
    activeConnection,
    companyName,
    realmId,
    lastSyncText,
    connectedCount,
    remainingCompanies,
    reloadConnections,
    handleAddCompanyClick,
    switchActiveCompany,
    handleDisconnectCompany,
    handleRenameCompany,
  };
}
