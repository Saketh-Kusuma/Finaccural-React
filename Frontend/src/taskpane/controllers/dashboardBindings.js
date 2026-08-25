/**
 * Event binding for the Dashboard view: header/account menu, detail blocks,
 * ERP connect buttons, the Change Company modal and the company list.
 *
 * Mixed into AppController as `bindDashboardView`. The data-action listeners
 * (Setup / Pull / Refresh / journal / tier) live in dataActionBindings.js and
 * are registered from the tail of this function so the overall binding order
 * matches the original single-function version exactly.
 */
import { AppState, getMaxCompaniesForPlan } from "../state/appState.js";
import { ViewRouter } from "../router/viewRouter.js";
import { ApiService } from "../services/apiService.js";
import { ExcelService } from "../services/excelService.js";
import { AuthService } from "../services/authService.js";
import { NotificationService } from "../services/notificationService.js";
import { DashboardService } from "../services/dashboardService.js";
import { bindDataActionHandlers } from "./dataActionBindings.js";
import { clearPullPageCursor } from "../batchDataLoader.js";

export function bindDashboardView() {

    // Notification bell / drawer / Clear All
    NotificationService.init();

    // Logouts
    const handleLogout = (e) => {
        if (e) e.preventDefault();
        AuthService.logout();
    };
    document.getElementById("btnBlockLogout")?.addEventListener("click", handleLogout);
    document.getElementById("btnDropdownLogout")?.addEventListener("click", handleLogout);
    document.getElementById("btnChangePlan")?.addEventListener("click", () => {
        ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));
        ViewRouter.show("Plans");
    });

    // Copy Subscription ID handler
    const handleCopySubId = () => {
        const subIdText = document.getElementById("dashSubId")?.textContent?.trim() || AppState.subscriptionId;
        if (subIdText) {
            navigator.clipboard.writeText(subIdText).then(() => {
                const btn = document.getElementById("btnCopySubId");
                if (btn) {
                    const originalHTML = btn.innerHTML;
                    btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>`;
                    setTimeout(() => {
                        btn.innerHTML = originalHTML;
                    }, 1800);
                }
            }).catch(err => {
                console.error("Copy failed: ", err);
            });
        }
    };
    document.getElementById("btnCopySubId")?.addEventListener("click", handleCopySubId);

    // Dropdown Toggle
    const toggleDropdown = (e) => {
        e.stopPropagation();
        // Close notification drawer first to prevent overlap
        const drawer = document.getElementById("notifDrawer");
        if (drawer) drawer.style.display = "none";

        const dropdown = document.getElementById("accountMenuDropdown");
        if (dropdown) {
            dropdown.style.display = dropdown.style.display === "none" ? "flex" : "none";
        }
    };
    document.getElementById("dashHeaderAvatarBtn1")?.addEventListener("click", toggleDropdown);
    document.getElementById("dashHeaderAvatarBtn2")?.addEventListener("click", toggleDropdown);
    document.getElementById("dashHeaderAvatarBtn3")?.addEventListener("click", toggleDropdown);
    document.getElementById("dashHeaderMenuBtn1")?.addEventListener("click", toggleDropdown);

    // Hide dropdown when clicking outside
    document.addEventListener("click", (e) => {
        const dropdown = document.getElementById("accountMenuDropdown");
        if (dropdown && !dropdown.contains(e.target)) {
            dropdown.style.display = "none";
        }
    });

    // Menu item to show block
    document.querySelectorAll(".dropdown-menu-item").forEach(item => {
        if (item.id === "btnDropdownLogout") return;
        item.addEventListener("click", (e) => {
            const targetId = e.currentTarget.dataset.target;
            document.querySelectorAll(".detail-block-card").forEach(c => c.style.display = "none");
            const targetBlock = document.getElementById(targetId);
            if (targetBlock) {
                targetBlock.style.display = "block";
            }
            const container = document.getElementById("detailBlocksContainer");
            if (container) {
                container.style.display = "flex";
            }
            const dropdown = document.getElementById("accountMenuDropdown");
            if (dropdown) dropdown.style.display = "none";
        });
    });

    // Close block buttons
    document.querySelectorAll(".close-block-btn, .close-card-btn").forEach(btn => {
        btn.addEventListener("click", (e) => {
            const card = e.target.closest(".detail-block-card");
            if (card) card.style.display = "none";
            const container = document.getElementById("detailBlocksContainer");
            if (container) {
                const anyVisible = Array.from(container.querySelectorAll(".detail-block-card")).some(c => c.style.display !== "none");
                if (!anyVisible) {
                    container.style.display = "none";
                }
            }
        });
    });

    // Close blocks on backdrop click
    const blocksContainer = document.getElementById("detailBlocksContainer");
    if (blocksContainer) {
        blocksContainer.addEventListener("click", (e) => {
            if (e.target === blocksContainer) {
                document.querySelectorAll(".detail-block-card").forEach(c => c.style.display = "none");
                blocksContainer.style.display = "none";
            }
        });
    }

    // Allow user to cancel connection attempt and go back to disconnected state
    document.getElementById("btnLogoutProvider")?.addEventListener("click", () => {
        DashboardService.renderERPSection();
    });

    // Connect QuickBooks — check if accounts already exist in DB first
    document.getElementById("btnConnectQB")?.addEventListener("click", async () => {
        try {
            const mail = AppState.userEmail || "";
            const res = await ApiService.apiFetch(`/api/connections?mail=${encodeURIComponent(mail)}`);
            const allConns = await res.json();
            const qbConns = (allConns || []).filter(c => (c.platform || "").toLowerCase() === "quickbooks");
            if (qbConns.length > 0) {
                // Existing QB accounts found — reactivate the first active one (or first overall)
                const toActivate = qbConns.find(c => c.status !== 'Disconnected') || qbConns[0];
                AppState.currentCompanyId = toActivate.companyId;
                AppState.currentProvider = "quickbooks";
                AppState.erpType = "quickbooks";
                await ApiService.apiFetch(`/api/connections/${toActivate.companyId}/activate`, { method: "POST" });
                DashboardService.renderERPSection();
                DashboardService.showStatus(`Resumed QuickBooks session for ${toActivate.companyName}`, "success", null, "quickbooks");
                return;
            }
        } catch (_) { }
        // No existing accounts — start OAuth
        DashboardService.showProviderSelected("quickbooks");
        DashboardService.launchERPOAuth("quickbooks");
    });

    // Connect Xero — check if accounts already exist in DB first
    document.getElementById("btnConnectXero")?.addEventListener("click", async () => {
        try {
            const mail = AppState.userEmail || "";
            const res = await ApiService.apiFetch(`/api/connections?mail=${encodeURIComponent(mail)}`);
            const allConns = await res.json();
            const xeroConns = (allConns || []).filter(c => (c.platform || "").toLowerCase() === "xero");
            if (xeroConns.length > 0) {
                // Existing Xero accounts found — reactivate the first one
                const toActivate = xeroConns.find(c => c.status !== 'Disconnected') || xeroConns[0];
                AppState.currentCompanyId = toActivate.companyId;
                AppState.currentProvider = "xero";
                AppState.erpType = "xero";
                await ApiService.apiFetch(`/api/connections/${toActivate.companyId}/activate`, { method: "POST" });
                DashboardService.renderERPSection();
                DashboardService.showStatus(`Resumed Xero session for ${toActivate.companyName}`, "success", null, "xero");
                return;
            }
        } catch (_) { }
        // No existing accounts — start OAuth
        DashboardService.showProviderSelected("xero");
        DashboardService.launchERPOAuth("xero");
    });

    // Connect Provider button in provider-selected state — launches OAuth
    document.getElementById("btnConnectProvider")?.addEventListener("click", () => {
        DashboardService.launchERPOAuth(AppState.currentProvider);
    });

    // Disconnect ERP (overall disconnect)
    document.getElementById("btnDisconnectERP")?.addEventListener("click", async () => {
        await DashboardService.disconnectERP();
    });

    // Sub card Change Plan button
    document.getElementById("btnSubChangePlan")?.addEventListener("click", () => {
        ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));
        ViewRouter.show("Plans");
    });

    // Add Company buttons (checks plan limits before starting OAuth)
    // Uses the currently active ERP provider so that:
    //   - Xero dashboard  → opens Xero OAuth
    //   - QuickBooks dashboard → opens QuickBooks OAuth
    const handleAddCompanyClick = () => {
        ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));
        const currentPlan = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null') ? AppState.subscriptionPlan : "Basic";
        const maxAllowed = getMaxCompaniesForPlan(currentPlan);
        // Derive the provider from the active dashboard, fall back to quickbooks
        const provider = AppState.currentProvider || "quickbooks";

        console.log("handleAddCompanyClick: provider = " + provider + ", AppState.currentProvider = " + AppState.currentProvider);

        ApiService.apiFetch("/api/connections?mail=" + encodeURIComponent(AppState.userEmail || ""))
            .then(r => r.json())
            .then(conns => {
                const connsForProvider = (conns || []).filter(c =>
                    (c.platform || "").toLowerCase() === provider.toLowerCase()
                );
                if (connsForProvider.length >= maxAllowed) {
                    alert(`Your ${currentPlan} Plan limits active ${provider === 'quickbooks' ? 'QuickBooks' : 'Xero'} connections to ${maxAllowed} companies. Please upgrade your plan to connect more companies.`);
                    ViewRouter.show("Plans");
                } else {
                    DashboardService.showProviderSelected(provider);
                    DashboardService.launchERPOAuth(provider);
                }
            })
            .catch(() => {
                DashboardService.launchERPOAuth(provider);
            });
    };
    document.getElementById("btnAddCompany")?.addEventListener("click", handleAddCompanyClick);
    document.getElementById("btnModalAddCompany")?.addEventListener("click", () => {
        const modal = document.getElementById("changeCompanyModal");
        if (modal) modal.style.display = "none";
        handleAddCompanyClick();
    });

    // Change Company Modal open/close & switch logic
    let selectedModalCompanyId = null;
    const openChangeCompanyModal = () => {
        const modal = document.getElementById("changeCompanyModal");
        const modalList = document.getElementById("modalCompanyList");
        if (!modal || !modalList) return;

        ApiService.apiFetch("/api/connections?mail=" + encodeURIComponent(AppState.userEmail || ""))
            .then(r => r.json())
            .then(conns => {
                modalList.innerHTML = "";
                // Filter to show only companies for the current platform
                const currentPlatform = AppState.currentProvider || "quickbooks";
                const platformConns = conns.filter(c => (c.platform || "quickbooks").toLowerCase() === currentPlatform);
                const platformLabel = currentPlatform === "xero" ? "Xero" : "QuickBooks";

                // Update modal title
                const modalTitle = document.querySelector("#changeCompanyModal .fa-modal-title, #changeCompanyModal h3");
                if (modalTitle) modalTitle.textContent = `Switch ${platformLabel} Company`;

                platformConns.forEach(c => {
                    const isSelected = c.companyId === (selectedModalCompanyId || AppState.currentCompanyId);
                    const isXero = (c.platform || "").toLowerCase() === "xero";
                    const isDisconnected = c.status === 'Disconnected';
                    const row = document.createElement("div");
                    row.className = `fa-modal-company-row ${isSelected ? "selected" : ""} ${isDisconnected ? "disconnected" : ""}`;
                    row.dataset.companyId = c.companyId;
                    row.dataset.platform = (c.platform || "quickbooks").toLowerCase();

                    const displayName = c.companyName || (isXero ? "Xero Organisation" : "QuickBooks Company");

                    row.innerHTML = `
                                <div class="fa-company-icon ${isXero ? 'xero-company-icon' : ''}">${isXero ? 'xero' : 'qb'}</div>
                                <div class="fa-company-info">
                                    <div class="fa-company-name">${displayName}${isDisconnected ? ' <span style="color:#ef4444;font-size:10px">(Disconnected)</span>' : ''}</div>
                                    <div class="fa-company-tag">${isXero ? "Tenant ID" : "Realm ID"}: ${c.companyId || "—"}</div>
                                </div>
                            `;
                    row.addEventListener("click", () => {
                        modalList.querySelectorAll(".fa-modal-company-row").forEach(r => r.classList.remove("selected"));
                        row.classList.add("selected");
                        selectedModalCompanyId = c.companyId;
                    });
                    modalList.appendChild(row);
                });
                modal.style.display = "flex";
            })
            .catch((err) => {
                // This chain had no .catch() before — an offline
                // backend (or any other apiFetch failure) rejected
                // silently as an unhandled promise rejection instead
                // of being caught here, which is what crashed to the
                // "Uncaught runtime errors" dev overlay. The modal
                // just doesn't open now, same as every other
                // apiFetch caller in this file already does on
                // failure.
                console.error("Error loading companies for Change Company modal:", err);
                DashboardService.showStatus(
                    "Couldn't load companies. Please check your connection and try again.",
                    "error",
                    null,
                    AppState.currentProvider
                );
            });
    };

    document.getElementById("btnChangeCompany")?.addEventListener("click", openChangeCompanyModal);
    document.getElementById("btnManageCompanies")?.addEventListener("click", openChangeCompanyModal);

    const closeChangeCompanyModal = () => {
        const modal = document.getElementById("changeCompanyModal");
        if (modal) modal.style.display = "none";
    };
    document.getElementById("btnCloseChangeCompany")?.addEventListener("click", closeChangeCompanyModal);
    document.getElementById("btnCancelChangeCompany")?.addEventListener("click", closeChangeCompanyModal);

    document.getElementById("btnConfirmChangeCompany")?.addEventListener("click", async () => {
        if (selectedModalCompanyId) {
            AppState.currentCompanyId = selectedModalCompanyId;
            // Determine the platform from the selected modal row
            const selectedRow = document.querySelector(`.fa-modal-company-row[data-company-id="${selectedModalCompanyId}"]`);
            if (selectedRow && selectedRow.dataset.platform) {
                AppState.currentProvider = selectedRow.dataset.platform;
                AppState.erpType = selectedRow.dataset.platform;
            }
            const switchedPlatform = AppState.currentProvider;
            ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));
            try {
                await ApiService.apiFetch(`/api/connections/${selectedModalCompanyId}/activate`, { method: "POST" });
            } catch (_) { }
            DashboardService.renderERPSection();
            DashboardService.showStatus("Active company switched successfully.", "success", null, switchedPlatform);
        }
        closeChangeCompanyModal();
    });

    // Modal search filter
    document.getElementById("companySearchInput")?.addEventListener("input", (e) => {
        const term = e.target.value.toLowerCase();
        document.querySelectorAll(".fa-modal-company-row").forEach(row => {
            const text = row.textContent.toLowerCase();
            row.style.display = text.includes(term) ? "flex" : "none";
        });
    });

    // Close context menu when clicking anywhere else
    document.addEventListener("click", (e) => {
        const menu = document.getElementById("companyContextMenu");
        if (menu && !menu.contains(e.target) && !e.target.classList.contains("fa-btn-dots")) {
            menu.style.display = "none";
        }
    });

    // Disconnect Active Company
    document.getElementById("btnDisconnectActiveCompany")?.addEventListener("click", async () => {
        const companyId = AppState.currentCompanyId;
        if (!companyId) return;
        const disconnectPlatform = AppState.currentProvider;

        DashboardService.showStatus("Disconnecting company...", "success", null, disconnectPlatform);
        try {
            const res = await ApiService.apiFetch(`/api/connections/${companyId}`, {
                method: "DELETE"
            });
            if (res.ok) {
                AppState.currentCompanyId = null; // Reset so a new active company gets picked
                DashboardService.showStatus("Company disconnected successfully.", "success", null, disconnectPlatform);
                DashboardService.renderERPSection();
            } else {
                DashboardService.showStatus("Failed to disconnect company.", "error", null, disconnectPlatform);
            }
        } catch (err) {
            DashboardService.showStatus("Error disconnecting company.", "error", null, disconnectPlatform);
        }
    });

    // Dropdown selection change
    document.getElementById("companySelectDropdown")?.addEventListener("change", (e) => {
        const dropdown = e.target;
        const opt = dropdown.options[dropdown.selectedIndex];
        if (opt) {
            AppState.currentCompanyId = opt.value;
            AppState.currentProvider = opt.dataset.platform;
            AppState.erpType = opt.dataset.platform;
            clearPullPageCursor(AppState.currentProvider, AppState.currentCompanyId);
            DashboardService.markStepIncomplete("setup");
            DashboardService.markStepIncomplete("pull");
            ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));

            // Activate in backend and re-render
            ApiService.apiFetch(`/api/connections/${opt.value}/activate`, { method: "POST" })
                .then(() => DashboardService.renderERPSection())
                .catch(() => DashboardService.renderERPSection());

            DashboardService.addLog(`Switched active company to: ${opt.textContent}`);
        }
    });


    // Setup Sheets / Pull Master Data / Refresh Schedule / journal / tier.
    bindDataActionHandlers();
}
