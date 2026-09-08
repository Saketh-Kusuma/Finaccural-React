import React, { useState } from "react";
import { initials } from "./ui";
import { AccountMenu } from "./AccountMenu";
import { openErp } from "../taskpane/api";
import { BellIcon } from "./dashboard/DashboardIcons";
import { useCompanyConnections } from "./dashboard/useCompanyConnections";
import { useMasterDataSync } from "./dashboard/useMasterDataSync";
import { CompanyListSection } from "./dashboard/CompanyListSection";
import { SubscriptionCard } from "./dashboard/SubscriptionCard";
import { JournalSelector } from "./dashboard/JournalSelector";
import { SyncActionsPanel } from "./dashboard/SyncActionsPanel";
import { ChangeCompanyModal } from "./dashboard/ChangeCompanyModal";
import { RenameCompanyModal } from "./dashboard/RenameCompanyModal";
import { CompanyContextMenu } from "./dashboard/CompanyContextMenu";

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
  const [showChangeModal, setShowChangeModal] = useState(false);
  const [showRenameModal, setShowRenameModal] = useState(false);
  const [renamingCompany, setRenamingCompany] = useState(null);
  const [contextMenu, setContextMenu] = useState({ visible: false, x: 0, y: 0, company: null });

  // Format plan for pill and stats
  const storedPlan =
    user.plan ||
    localStorage.getItem("fa_plan") ||
    localStorage.getItem("fa_subscription_plan") ||
    "Pro";
  const planClean = storedPlan.replace(/\s*plan\s*/i, "").trim() || "Pro";
  const planDisplay = `${planClean.toUpperCase()} PLAN`;
  const name = user.name || localStorage.getItem("fa_user_name") || "Sai";

  const companyConn = useCompanyConnections({
    provider,
    user,
    planClean,
    notify,
    onConnect,
    onChangePlan,
  });

  const syncState = useMasterDataSync({
    provider,
    activeConnection: companyConn.activeConnection,
    realmId: companyConn.realmId,
    planClean,
    label: companyConn.label,
    notify,
  });

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
      {/* ── HEADER ── */}
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
          Disconnect {companyConn.label}
        </button>
        <span className="fa-conn-dot">●</span>
        <span className="fa-conn-status">
          <span className="fa-realm-label">
            {companyConn.isXero ? "Tenant ID:" : "Realm ID:"}
          </span>{" "}
          <span>{companyConn.realmId}</span>
        </span>
      </div>

      {/* ── Account dropdown menu ── */}
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

      {/* ── SCROLLABLE BODY ── */}
      <div className="fa-body scrollable">
        <CompanyListSection
          label={companyConn.label}
          isXero={companyConn.isXero}
          provider={provider}
          platformConns={companyConn.platformConns}
          activeConnection={companyConn.activeConnection}
          companyName={companyConn.companyName}
          realmId={companyConn.realmId}
          lastSyncText={companyConn.lastSyncText}
          connectedCount={companyConn.connectedCount}
          onAddCompanyClick={companyConn.handleAddCompanyClick}
          onSwitchActiveCompany={companyConn.switchActiveCompany}
          onReconnectCompany={() => {
            notify("Launching re-authorization...", "success", null, provider);
            if (onConnect) onConnect(provider);
            else openErp(provider, user);
          }}
          onOpenContextMenu={(e, c) => {
            const rect = e.currentTarget.getBoundingClientRect();
            setContextMenu({
              visible: true,
              x: Math.max(10, rect.left - 100),
              y: rect.bottom + 4,
              company: c,
            });
          }}
        />

        <SubscriptionCard
          planClean={planClean}
          connectedCount={companyConn.connectedCount}
          maxCompanies={companyConn.maxCompanies}
          remainingCompanies={companyConn.remainingCompanies}
          onChangePlan={onChangePlan}
          onManageCompanies={() => setShowChangeModal(true)}
          notify={notify}
        />

        <JournalSelector
          journal={journal}
          onChangeJournal={setJournal}
        />

        <SyncActionsPanel
          label={companyConn.label}
          spinning={syncState.spinning}
          setupBusy={syncState.setupBusy}
          pullBusy={syncState.pullBusy}
          isSetupDone={syncState.isSetupDone}
          isPullDone={syncState.isPullDone}
          logs={syncState.logs}
          onRefresh={syncState.handleRefresh}
          onSetup={syncState.handleSetup}
          onPull={syncState.handlePull}
        />
      </div>

      {/* ── MODALS & MENUS ── */}
      <ChangeCompanyModal
        isOpen={showChangeModal}
        label={companyConn.label}
        platformConns={companyConn.platformConns}
        initialCompanyId={companyConn.activeConnection?.companyId || null}
        onClose={() => setShowChangeModal(false)}
        onSwitchCompany={companyConn.switchActiveCompany}
        onAddAnotherCompany={companyConn.handleAddCompanyClick}
      />

      <RenameCompanyModal
        isOpen={showRenameModal}
        company={renamingCompany}
        onClose={() => {
          setShowRenameModal(false);
          setRenamingCompany(null);
        }}
        onSave={(comp, newName) => companyConn.handleRenameCompany(comp, newName)}
      />

      <CompanyContextMenu
        visible={contextMenu.visible}
        x={contextMenu.x}
        y={contextMenu.y}
        company={contextMenu.company}
        onClose={() => setContextMenu({ visible: false, x: 0, y: 0, company: null })}
        onEdit={(comp) => {
          setRenamingCompany(comp);
          setShowRenameModal(true);
        }}
        onDisconnect={(comp) => companyConn.handleDisconnectCompany(comp)}
      />
    </section>
  );
}
