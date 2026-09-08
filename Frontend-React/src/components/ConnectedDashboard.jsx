import { useEffect, useState } from "react";
import { initials } from "./ui";
import { AccountMenu } from "./AccountMenu";
import {
  apiFetch,
  fetchMasterDataStream,
  fetchIncrementalDataStream,
  openErp,
} from "../taskpane/api";
import { ExcelService } from "../taskpane/services/excelService";
import { flattenAllMasterDataRecords } from "../taskpane/services/excelDataMappers";

function BellIcon() {
  return (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9" />
      <path d="M13.73 21a2 2 0 0 1-3.46 0" />
    </svg>
  );
}

function AddIcon() {
  return (
    <svg
      width="12"
      height="12"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.5"
      strokeLinecap="round"
    >
      <line x1="12" y1="5" x2="12" y2="19" />
      <line x1="5" y1="12" x2="19" y2="12" />
    </svg>
  );
}

function RefreshIcon({ spinning }) {
  return (
    <svg
      width="12"
      height="12"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.5"
      strokeLinecap="round"
      strokeLinejoin="round"
      className={`refresh-icon ${spinning ? "spin" : ""}`}
    >
      <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67" />
    </svg>
  );
}

function PlayIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
      <polygon points="5 3 19 12 5 21 5 3" />
    </svg>
  );
}

function DownloadIcon() {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 10 12 15 17 10" />
      <line x1="12" y1="15" x2="12" y2="3" />
    </svg>
  );
}

export function ConnectedDashboard({
  provider,
  user,
  disconnect,
  logout,
  onChangePlan,
  notify,
  unreadCount = 0,
  onToggleNotifications,
  onConnect,
}) {
  const [journal, setJournal] = useState("accrual");
  const [menu, setMenu] = useState(false);
  const [spinning, setSpinning] = useState(false);
  const [setupBusy, setSetupBusy] = useState(false);
  const [pullBusy, setPullBusy] = useState(false);
  const [connections, setConnections] = useState([]);
  const [activeCompanyId, setActiveCompanyId] = useState(
    () => localStorage.getItem("fa_current_company_id") || null
  );
  const [showChangeModal, setShowChangeModal] = useState(false);
  const [selectedModalCompanyId, setSelectedModalCompanyId] = useState(null);
  const [companySearchTerm, setCompanySearchTerm] = useState("");
  const [showRenameModal, setShowRenameModal] = useState(false);
  const [renamingCompany, setRenamingCompany] = useState(null);
  const [renameInputVal, setRenameInputVal] = useState("");
  const [contextMenu, setContextMenu] = useState({ visible: false, x: 0, y: 0, company: null });
  const [isSetupDone, setIsSetupDone] = useState(
    () => localStorage.getItem("fa_step_setup") === "complete"
  );
  const [isPullDone, setIsPullDone] = useState(false);
  const [logs, setLogs] = useState([]);

  const isXero = (provider || "").toLowerCase() === "xero";
  const label = isXero ? "Xero" : "QuickBooks";
  const name = user.name || localStorage.getItem("fa_user_name") || "Sai";

  const addLog = (msg) => {
    const time = new Date().toLocaleTimeString();
    setLogs((prev) => [...prev, `[${time}] ${msg}`]);
  };

  // Format plan for pill and stats
  const storedPlan =
    user.plan ||
    localStorage.getItem("fa_plan") ||
    localStorage.getItem("fa_subscription_plan") ||
    "Pro";
  const planClean = storedPlan.replace(/\s*plan\s*/i, "").trim() || "Pro";
  const planDisplay = `${planClean.toUpperCase()} PLAN`;

  // Plan limits: Pro = 10, Standard = 3, Basic = 1
  const maxCompanies = planClean.toLowerCase().includes("pro")
    ? 10
    : planClean.toLowerCase().includes("standard")
      ? 3
      : 1;

  // Format relative time matching vanilla formatRelativeTime
  const formatRelativeTime = (dateInput, status) => {
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
  };

  const reloadConnections = async () => {
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
  };

  // Fetch real connection details from backend
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

  // Listen for OAuth completion message from popup
  useEffect(() => {
    const receive = (event) => {
      if (event.data === "qb_connected" || event.data === "xero_connected") {
        addLog(`Connection completed: ${event.data}`);
        reloadConnections().then((conns) => {
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
  }, [provider]);

  // Platform connections
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

  const connectedCount = platformConns.length > 0 ? platformConns.length : (connections.length > 0 ? connections.length : 1);
  const remainingCompanies = Math.max(0, maxCompanies - connectedCount);

  // Exact Add Company Click Handler matching Vanilla frontend handleAddCompanyClick
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

    addLog(`Opening ${label} connection window...`);
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

  // Switch active company matching vanilla switchActiveCompany
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
        addLog(`Switched active company to: ${targetConn.companyName}${countMsg}`);
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

  // Disconnect company matching vanilla
  const handleDisconnectCompany = async (company) => {
    const compId = company.companyId;
    const compName = company.companyName || "Company";
    addLog(`Disconnecting ${compName}...`);
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
      addLog(`Company disconnected: ${compName}`);
      reloadConnections();
    } catch (_) {
      notify("Failed to disconnect company.", "error", null, provider);
    }
  };

  // Rename company matching vanilla
  const handleRenameCompany = async () => {
    const newName = renameInputVal.trim();
    if (!newName || !renamingCompany) return;
    try {
      await apiFetch(`/api/connections/${renamingCompany.companyId}/rename`, {
        method: "PATCH",
        body: JSON.stringify({ companyName: newName }),
      });
      notify("Company renamed successfully.", "success", null, provider);
      setShowRenameModal(false);
      setRenamingCompany(null);
      reloadConnections();
    } catch (_) {
      notify("Failed to rename company.", "error", null, provider);
    }
  };

  // Close context menu on outside click
  useEffect(() => {
    const handleClick = () => {
      if (contextMenu.visible) setContextMenu({ visible: false, x: 0, y: 0, company: null });
    };
    window.addEventListener("click", handleClick);
    return () => window.removeEventListener("click", handleClick);
  }, [contextMenu.visible]);

  // Setup Master & Input Sheets action
  const handleSetup = async () => {
    if (setupBusy) return;
    setSetupBusy(true);
    addLog(`Setting up Master & Input sheets for ${label}...`);
    notify("Setting up sheets...", "success", null, provider);
    try {
      await ExcelService.setupWorkbookSheets(provider);
      localStorage.setItem("fa_step_setup", "complete");
      setIsSetupDone(true);
      addLog("Sheets setup successfully.");
      notify("Master and Input sheets setup successfully.", "success", null, provider);
    } catch (err) {
      console.error(err);
      addLog("Error setting up sheets: " + (err.message || err));
      notify("Error setting up sheets.", "error", err.message || null, provider);
    } finally {
      setSetupBusy(false);
    }
  };

  // Pull Master Data from ERP action
  const handlePull = async () => {
    if (pullBusy) return;

    // Prerequisite check: exactly as in vanilla frontend!
    if (!isSetupDone) {
      const detailMsg = `Cannot pull master data: You must run Setup Master & Input Sheets for ${label} first.`;
      addLog(`Pull Master Data failed: ${detailMsg}`);
      notify("Pull Master Data Failed", "error", detailMsg, provider);
      return;
    }

    setPullBusy(true);
    const activeId = activeConnection?.companyId || realmId || "";
    addLog(`Pulling master data from ${label}...`);
    notify("Initializing Data Pull...", "success", "Pre-flight record count check in progress...", provider);

    try {
      const data = await fetchMasterDataStream(
        provider,
        activeId,
        planClean,
        (progress) => {
          if (progress.percentage) {
            notify(`Pulling Master Data (${progress.percentage}%)`, "success", `Fetched ${progress.fetchedRecords || 0} of ${progress.totalRecords || 0} records`, provider);
          }
        },
      );

      if (!data) {
        addLog("Data pull failed: No data returned.");
        notify("Data pull failed.", "error", "Please try again.", provider);
        return;
      }

      await ExcelService.clearMasterDataRange();
      const batch = flattenAllMasterDataRecords(data, { includeCompany: true });

      if (batch.length === 0) {
        addLog("Pull: no master data records found for this company.");
        notify("No more data available.", "success", "No master data found for this company.", provider);
        setIsPullDone(true);
        return;
      }

      const count = await ExcelService.appendManualBatch(provider, batch);
      setIsPullDone(true);
      const pullTitle = "Data completed.";
      const pullDetail = `Successfully fetched all ${count} records across all entities for this company.`;
      addLog(`${pullTitle} (${count} records pulled)`);
      notify(pullTitle, "success", pullDetail, provider);
    } catch (err) {
      console.error(err);
      addLog("Error pulling data: " + (err.message || err));
      notify("Data pull failed.", "error", err.message || "Please try again.", provider);
    } finally {
      setPullBusy(false);
    }
  };

  // Refresh Schedule action
  const handleRefresh = async () => {
    if (spinning) return;
    setSpinning(true);
    const activeId = activeConnection?.companyId || realmId || "";
    addLog(`Refreshing live data from ${label}...`);
    notify("Refreshing...", "success", null, provider);
    try {
      const data = await fetchIncrementalDataStream(
        provider,
        activeId,
        planClean,
      );
      if (data) {
        const batch = flattenAllMasterDataRecords(data, {
          includeCompany: false,
        });
        const updatedCount = await ExcelService.appendManualBatch(
          provider,
          batch,
        );
        const timestamp = new Date().toLocaleTimeString();
        await ExcelService.stampLastRefreshed(timestamp);
        if (updatedCount === 0) {
          addLog("Schedule Refreshed: No new records found.");
          notify("Schedule Refreshed", "success", "No new Records Found .", provider);
        } else {
          addLog(`Schedule Refreshed: ${updatedCount} records added.`);
          notify(
            "Schedule Refreshed",
            "success",
            `${updatedCount} updated record${updatedCount === 1 ? "" : "s"} added.`,
            provider
          );
        }
      }
    } catch (err) {
      console.error(err);
      addLog("Error refreshing data: " + (err.message || err));
      notify("Data refresh failed.", "error", "Please try again.", provider);
    } finally {
      setSpinning(false);
    }
  };

  return (
    <section
      className="view active"
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        position: "relative",
      }}
    >
      {/* ── PREMIUM HEADER (Matches Vanilla fa-header) ── */}
      <div className="fa-header">
        <div className="fa-header-left">
          <div className="fa-logo-box">FA</div>
          <div className="fa-header-text">
            <div className="fa-brand-name">FinAccrual</div>
          </div>
        </div>
        <div className="fa-header-right">
          <span className="fa-plan-badge" id="connTierBadge">
            {planDisplay}
          </span>
          <button
            className="fa-notif-bell-btn"
            title="Notifications"
            aria-label="Notifications"
            onClick={() => onToggleNotifications && onToggleNotifications("88px")}
          >
            <BellIcon />
            {unreadCount > 0 && (
              <span className="fa-notif-badge">
                {unreadCount > 99 ? "99+" : unreadCount}
              </span>
            )}
          </button>
          <button
            className="fa-avatar-btn"
            onClick={() => setMenu(!menu)}
            aria-label="Account menu"
          >
            {initials(name)}
          </button>
        </div>
      </div>

      {/* ── Header row 2: Realm/Tenant + status ── */}
      <div className="fa-header-row2">
        <button className="fa-disconnect-btn" onClick={disconnect}>
          Disconnect {label}
        </button>
        <span className="fa-conn-dot">●</span>
        <span className="fa-conn-status">
          <span className="fa-realm-label">
            {isXero ? "Tenant ID:" : "Realm ID:"}
          </span>{" "}
          <span>{realmId}</span>
        </span>
      </div>

      {/* ── Account dropdown menu (Exact Vanilla Frontend Menu with Icons) ── */}
      {menu && (
        <AccountMenu
          user={user}
          planClean={planClean}
          onClose={() => setMenu(false)}
          onLogout={logout}
          onChangePlan={onChangePlan || (() => notify("Plan selector opened.", "success"))}
          notify={notify}
        />
      )}

      {/* ── SCROLLABLE BODY (Matches Vanilla fa-body scrollable) ── */}
      <div className="fa-body scrollable">
        {/* COMPANY MANAGEMENT SECTION */}
        <div className="fa-section-header">
          <span className="fa-section-title">
            {label.toUpperCase()} COMPANIES
          </span>
          <button
            className="fa-btn-add"
            onClick={handleAddCompanyClick}
          >
            <AddIcon />
            Add Another Company
          </button>
        </div>

        {/* Company List Card */}
        <div className="fa-company-list">
          {platformConns.length === 0 ? (
            <div className="fa-company-item active-company">
              <input
                type="radio"
                name="companyRadio"
                className="fa-company-radio"
                checked
                readOnly
              />
              <div
                className={`fa-company-icon ${isXero ? "xero-company-icon" : ""}`}
              >
                {isXero ? "xero" : "qb"}
              </div>
              <div className="fa-company-info">
                <div className="fa-company-name">{companyName}</div>
                <div className="fa-company-tag">Last Sync: {lastSyncText}</div>
              </div>
              <div className="fa-company-actions">
                <span className="fa-badge-active">ACTIVE</span>
                <button
                  className="fa-btn-dots"
                  title="More options"
                  onClick={(e) => {
                    e.stopPropagation();
                    const rect = e.currentTarget.getBoundingClientRect();
                    setContextMenu({
                      visible: true,
                      x: Math.max(10, rect.left - 100),
                      y: rect.bottom + 4,
                      company: { companyId: realmId, companyName, platform: provider },
                    });
                  }}
                >
                  ⋮
                </button>
              </div>
            </div>
          ) : (
            platformConns.map((c) => {
              const isActive = c.companyId === activeConnection?.companyId;
              const isDisconnected = c.status === "Disconnected";
              const cIsXero = (c.platform || "").toLowerCase() === "xero";
              const cDisplayName = c.companyName || (cIsXero ? "Xero Organisation" : "QuickBooks Company");
              const cLastSync = formatRelativeTime(c.lastSyncedAt, c.status);

              return (
                <div
                  key={c.companyId}
                  className={`fa-company-item ${isActive && !isDisconnected ? "active-company" : ""} ${isDisconnected ? "disconnected-company" : ""}`}
                  style={{ cursor: isDisconnected ? "default" : "pointer" }}
                  onClick={() => {
                    if (!isDisconnected && !isActive) {
                      switchActiveCompany(c.companyId);
                    }
                  }}
                >
                  <input
                    type="radio"
                    name="companyRadio"
                    className="fa-company-radio"
                    checked={isActive && !isDisconnected}
                    onChange={() => {
                      if (!isDisconnected) switchActiveCompany(c.companyId);
                    }}
                  />
                  <div
                    className={`fa-company-icon ${cIsXero ? "xero-company-icon" : ""}`}
                  >
                    {cIsXero ? "xero" : "qb"}
                  </div>
                  <div className="fa-company-info">
                    <div className="fa-company-name">
                      {cDisplayName}
                      {isDisconnected && (
                        <span style={{ color: "#ef4444", fontSize: 10, marginLeft: 4 }}>(Disconnected)</span>
                      )}
                    </div>
                    <div className="fa-company-tag">Last Sync: {cLastSync}</div>
                  </div>
                  <div className="fa-company-actions">
                    {isActive && !isDisconnected ? (
                      <span className="fa-badge-active">ACTIVE</span>
                    ) : isDisconnected ? (
                      <button
                        className="fa-btn-reconnect"
                        onClick={(e) => {
                          e.stopPropagation();
                          notify("Launching re-authorization...", "success", null, provider);
                          if (onConnect) onConnect(provider);
                          else openErp(provider, user);
                        }}
                      >
                        Reconnect
                      </button>
                    ) : (
                      <button
                        className="fa-btn-switch"
                        onClick={(e) => {
                          e.stopPropagation();
                          switchActiveCompany(c.companyId);
                        }}
                      >
                        Switch
                      </button>
                    )}
                    <button
                      className="fa-btn-dots"
                      title="More options"
                      onClick={(e) => {
                        e.stopPropagation();
                        const rect = e.currentTarget.getBoundingClientRect();
                        setContextMenu({
                          visible: true,
                          x: Math.max(10, rect.left - 100),
                          y: rect.bottom + 4,
                          company: c,
                        });
                      }}
                    >
                      ⋮
                    </button>
                  </div>
                </div>
              );
            })
          )}
        </div>
        <div className="fa-company-list-footer">
          Showing {connectedCount} {label}{" "}
          {connectedCount === 1 ? "company" : "companies"}
        </div>

        {/* SUBSCRIPTION INFO CARD */}
        <div className="fa-sub-card">
          <div className="fa-sub-title">CURRENT PLAN</div>
          <div className="fa-sub-stats">
            <div className="fa-sub-stat">
              <div className="fa-sub-stat-label">CURRENT PLAN</div>
              <div className="fa-sub-stat-value">{planClean}</div>
            </div>
            <div className="fa-sub-stat-divider"></div>
            <div className="fa-sub-stat">
              <div className="fa-sub-stat-label">CONNECTED COMPANIES</div>
              <div className="fa-sub-stat-value">
                {connectedCount} / {maxCompanies}
              </div>
            </div>
            <div className="fa-sub-stat-divider"></div>
            <div className="fa-sub-stat">
              <div className="fa-sub-stat-label">REMAINING</div>
              <div className="fa-sub-stat-value">{remainingCompanies}</div>
            </div>
          </div>
          <div className="fa-sub-actions">
            <button
              className="fa-btn-sub-outline"
              onClick={onChangePlan || (() => notify("Plan selector opened.", "success"))}
            >
              Change Plan
            </button>
            <button
              className="fa-btn-sub-outline"
              onClick={() => {
                setSelectedModalCompanyId(activeConnection?.companyId || null);
                setCompanySearchTerm("");
                setShowChangeModal(true);
              }}
            >
              Manage Companies
            </button>
          </div>
        </div>

        {/* JOURNAL TYPE SECTION */}
        <div className="fa-section-label">JOURNAL TYPE</div>
        <div className="fa-journal-btns">
          <button
            className={`fa-journal-btn ${journal === "accrual" ? "active" : ""}`}
            onClick={() => setJournal("accrual")}
          >
            Accrual
          </button>
          <button
            className={`fa-journal-btn ${journal === "prepaid" ? "active" : ""}`}
            onClick={() => setJournal("prepaid")}
          >
            Prepaid
          </button>
          <button
            className={`fa-journal-btn ${journal === "deferred" ? "active" : ""}`}
            onClick={() => setJournal("deferred")}
          >
            Deferred Rev
          </button>
        </div>

        {/* SCHEDULE SECTION */}
        <div className="schedule-header-row">
          <span className="fa-section-label">SCHEDULE</span>
          <button
            className="refresh-schedule-btn"
            onClick={handleRefresh}
            disabled={spinning}
          >
            <RefreshIcon spinning={spinning} />
            {spinning ? "Refreshing…" : "Refresh Schedule"}
          </button>
        </div>

        <button
          className="conn-action-btn primary-btn"
          onClick={handleSetup}
          disabled={setupBusy}
        >
          <PlayIcon />
          {setupBusy ? "Setting up sheets…" : "Setup Master & Input Sheets"}
        </button>

        <button
          className="conn-action-btn secondary-btn"
          onClick={handlePull}
          disabled={pullBusy}
        >
          <DownloadIcon />
          {pullBusy ? "Pulling Master Data…" : `Pull Master Data from ${label}`}
        </button>

        {/* PROGRESS CONSOLE (Exact match to vanilla frontend) */}
        <div className="conn-progress-panel">
          <div className="conn-step complete">
            <span className="conn-step-num">1</span>
            <span className="conn-step-label">Connect to {label}</span>
          </div>
          <div className={`conn-step ${isSetupDone ? "complete" : setupBusy ? "active" : ""}`}>
            <span className="conn-step-num">2</span>
            <span className="conn-step-label">Setup Master &amp; Input Sheets</span>
          </div>
          <div className={`conn-step ${isPullDone ? "complete" : pullBusy ? "active" : ""}`}>
            <span className="conn-step-num">3</span>
            <span className="conn-step-label">Pull Master Data from {label}</span>
          </div>
          <div className="conn-log">
            {logs.map((line, idx) => {
              const isError = /error|fail|failed|unable|denied|invalid/i.test(line);
              return (
                <div key={idx} className={`log-line ${isError ? "log-error" : ""}`}>
                  {line}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* ── CHANGE COMPANY MODAL (Exact match to Vanilla fa-modal-overlay) ── */}
      {showChangeModal && (
        <div className="fa-modal-overlay" onClick={() => setShowChangeModal(false)}>
          <div className="fa-modal" onClick={(e) => e.stopPropagation()}>
            <div className="fa-modal-header">
              <span className="fa-modal-title">Switch {label} Company</span>
              <button
                className="fa-modal-close"
                onClick={() => setShowChangeModal(false)}
              >
                ✕
              </button>
            </div>
            <div className="fa-modal-search-wrap">
              <svg
                width="14"
                height="14"
                viewBox="0 0 24 24"
                fill="none"
                stroke="#6b7a9a"
                strokeWidth="2"
                strokeLinecap="round"
              >
                <circle cx="11" cy="11" r="8" />
                <line x1="21" y1="21" x2="16.65" y2="16.65" />
              </svg>
              <input
                type="text"
                className="fa-modal-search"
                placeholder="Search companies..."
                value={companySearchTerm}
                onChange={(e) => setCompanySearchTerm(e.target.value)}
              />
            </div>
            <div className="fa-modal-list">
              {platformConns
                .filter((c) =>
                  (c.companyName || "").toLowerCase().includes(companySearchTerm.toLowerCase())
                )
                .map((c) => {
                  const isSelected = c.companyId === selectedModalCompanyId;
                  const cIsXero = (c.platform || "").toLowerCase() === "xero";
                  const isDisconnected = c.status === "Disconnected";
                  const displayName = c.companyName || (cIsXero ? "Xero Organisation" : "QuickBooks Company");

                  return (
                    <div
                      key={c.companyId}
                      className={`fa-modal-company-row ${isSelected ? "selected" : ""} ${isDisconnected ? "disconnected" : ""}`}
                      onClick={() => setSelectedModalCompanyId(c.companyId)}
                    >
                      <div className={`fa-company-icon ${cIsXero ? "xero-company-icon" : ""}`}>
                        {cIsXero ? "xero" : "qb"}
                      </div>
                      <div className="fa-company-info">
                        <div className="fa-company-name">
                          {displayName}
                          {isDisconnected && (
                            <span style={{ color: "#ef4444", fontSize: 10, marginLeft: 4 }}>(Disconnected)</span>
                          )}
                        </div>
                        <div className="fa-company-tag">
                          {cIsXero ? "Tenant ID" : "Realm ID"}: {c.companyId || "—"}
                        </div>
                      </div>
                    </div>
                  );
                })}
            </div>
            <div className="fa-modal-link-row">
              <span className="fa-modal-link-text">Can't find your company?</span>
              <button
                className="fa-modal-add-link"
                onClick={() => {
                  setShowChangeModal(false);
                  handleAddCompanyClick();
                }}
              >
                + Add Another Company
              </button>
            </div>
            <div className="fa-modal-footer">
              <button
                className="fa-btn-modal-cancel"
                onClick={() => setShowChangeModal(false)}
              >
                Cancel
              </button>
              <button
                className="fa-btn-modal-primary"
                onClick={() => {
                  if (selectedModalCompanyId) {
                    switchActiveCompany(selectedModalCompanyId);
                  }
                  setShowChangeModal(false);
                }}
              >
                Switch Company
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── COMPANY THREE-DOT CONTEXT MENU ── */}
      {contextMenu.visible && (
        <div
          className="fa-context-menu"
          style={{
            display: "block",
            top: contextMenu.y,
            left: contextMenu.x,
          }}
          onClick={(e) => e.stopPropagation()}
        >
          <button
            className="fa-context-item"
            onClick={() => {
              const comp = contextMenu.company;
              setContextMenu({ visible: false, x: 0, y: 0, company: null });
              if (comp) {
                setRenamingCompany(comp);
                setRenameInputVal(comp.companyName || "");
                setShowRenameModal(true);
              }
            }}
          >
            Edit
          </button>
          <button
            className="fa-context-item"
            onClick={() => {
              const comp = contextMenu.company;
              setContextMenu({ visible: false, x: 0, y: 0, company: null });
              if (comp) handleDisconnectCompany(comp);
            }}
          >
            Disconnect
          </button>
        </div>
      )}

      {/* ── RENAME COMPANY MODAL ── */}
      {showRenameModal && (
        <div className="fa-modal-overlay" onClick={() => setShowRenameModal(false)}>
          <div className="fa-modal" onClick={(e) => e.stopPropagation()}>
            <div className="fa-modal-header">
              <span className="fa-modal-title">Rename Company</span>
              <button
                className="fa-modal-close"
                onClick={() => setShowRenameModal(false)}
              >
                ✕
              </button>
            </div>
            <div style={{ padding: 16 }}>
              <label
                style={{
                  display: "block",
                  fontSize: 12,
                  fontWeight: 600,
                  marginBottom: 6,
                  color: "#4a5568",
                }}
              >
                Company Display Name
              </label>
              <input
                type="text"
                className="fa-modal-search"
                style={{
                  width: "100%",
                  boxSizing: "border-box",
                  border: "1.5px solid #e2e8f4",
                  borderRadius: 8,
                  padding: "8px 10px",
                }}
                value={renameInputVal}
                onChange={(e) => setRenameInputVal(e.target.value)}
                autoFocus
              />
            </div>
            <div className="fa-modal-footer">
              <button
                className="fa-btn-modal-cancel"
                onClick={() => setShowRenameModal(false)}
              >
                Cancel
              </button>
              <button
                className="fa-btn-modal-primary"
                onClick={handleRenameCompany}
              >
                Save
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
