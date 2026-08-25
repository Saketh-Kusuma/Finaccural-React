/**
 * Dashboard UI: rendering, ERP connection status and company management.
 *
 * DashboardService is assembled from three parts that share one `this`:
 *   - this file          : render/status/company-list UI
 *   - dashboardConsole.js: activity log + progress steps
 *   - erpConnectionService.js: OAuth launch, connect/disconnect
 * They are merged (not nested) so every existing `DashboardService.x()` call
 * site keeps working unchanged.
 */
import { AppState, getMaxCompaniesForPlan } from "../state/appState.js";
import { ApiService } from "./apiService.js";
import { ExcelService } from "./excelService.js";
import { NotificationService } from "./notificationService.js";
import { createDashboardConsole } from "./dashboardConsole.js";
import { createErpConnection } from "./erpConnectionService.js";

const DashboardService = {

    /**
     * Renders and populates all dashboard UI elements based on AppState.
     */
    render() {
        const name = AppState.userName || AppState.userEmail || "User";
        const first = name.split(" ")[0];
        const initial = name.charAt(0).toUpperCase();

        // Set avatar initials
        const avatars = ["dashHeaderAvatarBtn1", "dashHeaderAvatarBtn2", "dashHeaderAvatarBtn3", "dropdownAvatar", "blockAvatar"];
        avatars.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.textContent = initial;
        });

        // Render details in disconnected state header
        const welcomeEl = document.getElementById("dashWelcome");
        const badgeEl = document.getElementById("dashPlanBadge");
        const subIdEl = document.getElementById("dashSubId");

        if (welcomeEl) welcomeEl.textContent = `Welcome, ${first}!`;
        if (badgeEl) badgeEl.textContent = ((AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null' ? AppState.subscriptionPlan : "Basic") + " Plan");
        if (subIdEl) subIdEl.textContent = AppState.subscriptionId || "—";

        // Render details in dropdown and blocks
        if (document.getElementById("dropdownUserName")) document.getElementById("dropdownUserName").textContent = name;
        if (document.getElementById("dropdownUserEmail")) document.getElementById("dropdownUserEmail").textContent = AppState.userEmail;
        if (document.getElementById("blockUserName")) document.getElementById("blockUserName").textContent = name;
        if (document.getElementById("blockUserEmail")) document.getElementById("blockUserEmail").textContent = AppState.userEmail;

        if (document.getElementById("blockSubId")) document.getElementById("blockSubId").textContent = AppState.subscriptionId || "—";
        const _safePlan = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null') ? AppState.subscriptionPlan : "Basic";
        if (document.getElementById("blockPlanName")) document.getElementById("blockPlanName").textContent = _safePlan + " Plan";
        if (document.getElementById("blockPlanTitle")) document.getElementById("blockPlanTitle").textContent = _safePlan + " Plan";
        if (document.getElementById("blockPlanType")) document.getElementById("blockPlanType").textContent = _safePlan + " Plan";

        // Render correct state sections
        this.renderERPSection();

        // Set connected values if connected
        if (AppState.erpConnected && AppState.erpType) {
            AppState.currentProvider = AppState.erpType;
            this.renderERPConsole();
        }
    },

    /**
     * Shows the provider-selected "not connected" state (Image 4).
     * Called when user clicks QuickBooks or Xero card from the connect screen.
     * @param {"quickbooks"|"xero"} provider
     */
    showProviderSelected(provider) {
        AppState.currentProvider = provider;
        const isQB = provider === "quickbooks";
        const pName = isQB ? "QuickBooks" : "Xero";

        // Hide disconnected + connected, show provider-selected
        const discSection = document.getElementById("dashDisconnected");
        const provSection = document.getElementById("dashProviderSelected");
        const connSection = document.getElementById("dashConnected");
        const connectingSection = document.getElementById("dashConnecting");
        if (discSection) discSection.style.display = "none";
        if (provSection) provSection.style.display = "flex";
        if (connSection) connSection.style.display = "none";
        if (connectingSection) connectingSection.style.display = "none";

        // Update the plan badge
        const planBadge = document.getElementById("providerPlanBadge");
        if (planBadge) {
            planBadge.textContent = isQB ? "QBO PRO" : "XERO PRO";
        }

        // Update the connect button
        const connectBtn = document.getElementById("btnConnectProvider");
        if (connectBtn) connectBtn.textContent = `Connect ${pName}`;

        // Update pull button label
        const pullLabel = document.getElementById("pullBtnProvLabel");
        if (pullLabel) pullLabel.textContent = isQB ? "QBO" : "Xero";

        // Update progress step labels
        const step1Label = document.getElementById("provStep1Label");
        const step3Label = document.getElementById("provStep3Label");
        if (step1Label) step1Label.textContent = `Connect to ${pName}`;
        if (step3Label) step3Label.textContent = `Pull Master Data`;

        // Active provider just changed — keep the badge/drawer, log
        // console, and step indicators in sync (all re-scoped to this
        // provider + whatever company is active for it).
        NotificationService.refreshForContext();
        this.renderActiveLogConsole();
        this.applyStepState();
    },

    showConnecting(provider) {
        AppState.currentProvider = provider;
        const isQB = provider === "quickbooks";
        const brandName = isQB ? "Intuit" : "Xero";

        const discSection = document.getElementById("dashDisconnected");
        const provSection = document.getElementById("dashProviderSelected");
        const connSection = document.getElementById("dashConnected");
        const connectingSection = document.getElementById("dashConnecting");

        if (discSection) discSection.style.display = "none";
        if (provSection) provSection.style.display = "none";
        if (connSection) connSection.style.display = "none";
        if (connectingSection) {
            connectingSection.style.display = "flex";

            // Reset the "Redirecting to <provider>..." card to its
            // initial state every time it's shown — provider branding,
            // full 5-second countdown, empty progress bar, no dots lit.
            // launchERPOAuth() drives the countdown/progress/dots from
            // here on; this just establishes the starting state.
            const card = document.getElementById("redirectCard");
            if (card) card.dataset.provider = isQB ? "qb" : "xero";

            const logoTextEl = document.getElementById("redirectLogoText");
            if (logoTextEl) logoTextEl.textContent = isQB ? "qb" : "xero";

            const brandNameEl = document.getElementById("redirectBrandName");
            if (brandNameEl) brandNameEl.textContent = brandName.toUpperCase();

            const textEl = document.getElementById("connectingText");
            if (textEl) textEl.textContent = `Redirecting to ${brandName}...`;

            const subtextEl = document.getElementById("redirectSubtext");
            if (subtextEl) subtextEl.textContent = `Securely connecting to your ${brandName} account`;

            const progressFillEl = document.getElementById("redirectProgressFill");
            if (progressFillEl) progressFillEl.style.width = "0%";

            const countdownEl = document.getElementById("redirectCountdown");
            if (countdownEl) countdownEl.textContent = "Please wait, opening in 5 seconds...";

            document.querySelectorAll("#redirectDots .dot").forEach(dot => dot.classList.remove("active"));

            // A previous attempt may have left the card on the Phase 2
            // "waiting" view (see showRedirectWaitingState below) — a
            // fresh Connect click always restarts on Phase 1.
            const mainView = document.getElementById("redirectMainView");
            const waitingView = document.getElementById("redirectWaitingView");
            if (mainView) mainView.style.display = "";
            if (waitingView) waitingView.style.display = "none";
        }

        // Active provider just changed — keep the badge/drawer, log
        // console, and step indicators in sync.
        NotificationService.refreshForContext();
        this.renderActiveLogConsole();
        this.applyStepState();
    },

    /**
     * Phase 2 of the redirect card (#dashConnecting): once the
     * QuickBooks/Xero sign-in window or dialog is confirmed open,
     * launchERPOAuth() calls this to swap the 5-second countdown for a
     * lightweight "Waiting for you to sign in..." spinner — the popup
     * is now the user's focus, so the task pane doesn't need to keep
     * showing an "opening in N seconds" countdown that already hit
     * zero, and it's too soon to drop to the full interactive
     * "provider selected" dashboard since nothing is connected yet.
     * #dashConnecting stays visible throughout; renderERPSection()
     * (reached via onERPConnected/cancelERPConnection once the OAuth
     * flow actually finishes) is what hides it and shows the real
     * final state.
     * @param {"quickbooks"|"xero"} provider
     */
    showRedirectWaitingState(provider) {
        const isQB = provider === "quickbooks";
        const brandName = isQB ? "Intuit" : "Xero";

        const mainView = document.getElementById("redirectMainView");
        const waitingView = document.getElementById("redirectWaitingView");
        if (mainView) mainView.style.display = "none";
        if (waitingView) waitingView.style.display = "block";

        const subtextEl = document.getElementById("redirectWaitingSubtext");
        if (subtextEl) {
            subtextEl.textContent = `Complete sign-in with ${brandName} in the window that opened — this will update automatically.`;
        }
    },

    /**
     * Shows the correct ERP section based on state:
     * - disconnected: connect cards (Image 2)
     * - provider-selected: not connected (Image 4)
     * - connected: fully connected (Image 3)
     */
    renderERPSection() {
        const discSection = document.getElementById("dashDisconnected");
        const provSection = document.getElementById("dashProviderSelected");
        const connSection = document.getElementById("dashConnected");
        const connectingSection = document.getElementById("dashConnecting");

        ApiService.apiFetch("/api/connections?mail=" + encodeURIComponent(AppState.userEmail || ""))
            .then(r => r.json())
            .then(conns => {
                const dropdown = document.getElementById("companySelectDropdown");
                const companyListEl = document.getElementById("companyList");
                const footerEl = document.getElementById("companyListFooter");
                const modalListEl = document.getElementById("modalCompanyList");

                const activeConns = conns.filter(c => c.status !== 'Disconnected');
                if (activeConns.length > 0) {
                    AppState.forceWelcome = false;
                }
                if (!conns || conns.length === 0 || AppState.forceWelcome) {
                    AppState.erpConnected = false;
                    if (discSection) {
                        discSection.style.display = "flex";
                        discSection.style.flexDirection = "column";
                        discSection.style.height = "100%";
                    }
                    if (provSection) provSection.style.display = "none";
                    if (connSection) connSection.style.display = "none";
                    if (connectingSection) connectingSection.style.display = "none";

                    // Update QB card button label
                    const hasQB = (conns || []).some(c => (c.platform || "").toLowerCase() === "quickbooks");
                    const hasXero = (conns || []).some(c => (c.platform || "").toLowerCase() === "xero");
                    const qbBtn = document.querySelector("#btnConnectQB .btn-connect-full");
                    const xeroBtn = document.querySelector("#btnConnectXero .btn-connect-full");
                    if (qbBtn) qbBtn.innerHTML = hasQB ? " Connect QuickBooks →" : `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg> Connect QuickBooks →`;
                    if (xeroBtn) xeroBtn.innerHTML = hasXero ? " Connect Xero  →" : `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg> Connect Xero →`;

                    // No active provider anymore — recompute the badge/drawer,
                    // log console, and step indicators so nothing left over
                    // from a previous provider/company lingers.
                    NotificationService.refreshForContext();
                    this.renderActiveLogConsole();
                    this.applyStepState();
                    return;
                }

                // We have connections!
                AppState.erpConnected = true;
                if (discSection) discSection.style.display = "none";
                if (provSection) provSection.style.display = "none";
                if (connSection) connSection.style.display = "flex";
                if (connectingSection) connectingSection.style.display = "none";

                // Determine the current platform first (respect AppState.currentProvider if already set)
                let resolvedProvider = AppState.currentProvider || null;
                if (!resolvedProvider) {
                    // Derive from the currently tracked companyId if possible
                    const existingConn = activeConns.find(c => c.companyId === AppState.currentCompanyId) || conns.find(c => c.companyId === AppState.currentCompanyId) || conns[0];
                    resolvedProvider = existingConn ? (existingConn.platform || "quickbooks").toLowerCase() : "quickbooks";
                }
                AppState.currentProvider = resolvedProvider;
                AppState.erpType = resolvedProvider;

                // Get active connections scoped to the current platform
                const platformActiveConns = activeConns.filter(c => (c.platform || "quickbooks").toLowerCase() === resolvedProvider);
                const fallbackConns = platformActiveConns.length > 0 ? platformActiveConns : conns.filter(c => (c.platform || "quickbooks").toLowerCase() === resolvedProvider);

                // Ensure current active company ID is valid for this platform
                const isCurrentValid = AppState.currentCompanyId && platformActiveConns.some(ac => ac.companyId === AppState.currentCompanyId);
                if (!isCurrentValid && fallbackConns.length > 0) {
                    AppState.currentCompanyId = fallbackConns[0].companyId;
                    // Update provider if we fell back to a different platform
                    const fb = fallbackConns[0];
                    AppState.currentProvider = (fb.platform || "quickbooks").toLowerCase();
                    AppState.erpType = AppState.currentProvider;
                }

                // Determine provider
                const isXero = (AppState.currentProvider || "quickbooks").toLowerCase() === "xero";

                // Handle Section Visibility (Xero Multi-select vs QB Single Active Company)
                const selectXeroCard = document.getElementById("selectXeroCompaniesCard");
                const activeCompanyCard = document.getElementById("activeCompanyCard");

                if (isXero) {
                    if (selectXeroCard) selectXeroCard.style.display = "block";
                    if (activeCompanyCard) activeCompanyCard.style.display = "none";

                    // Populate Xero Multi-Select Checklist Card
                    const currentPlan = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null') ? AppState.subscriptionPlan : "Basic";
                    const maxAllowed = getMaxCompaniesForPlan(currentPlan);
                    const xeroConns = conns.filter(c => (c.platform || "").toLowerCase() === "xero");

                    let xeroSelected = new Set(xeroConns.filter(c => c.status !== 'Disconnected').map(c => c.companyId));

                    const listEl = document.getElementById("xeroCompanyList");
                    const selCountEl = document.getElementById("xeroSelCount");
                    const maxCountEl = document.getElementById("xeroMaxCount");
                    const maxAllowedEl = document.getElementById("xeroMaxAllowed");
                    const planBadgeEl = document.getElementById("xeroPlanBadge");
                    const warningEl = document.getElementById("xeroLimitWarning");
                    const confirmBtn = document.getElementById("btnConfirmXeroCompanies");

                    if (maxCountEl) maxCountEl.textContent = maxAllowed;
                    if (maxAllowedEl) maxAllowedEl.textContent = maxAllowed;
                    if (planBadgeEl) planBadgeEl.textContent = `${currentPlan.toUpperCase()} (${maxAllowed}) PLAN`;

                    const updateXeroUI = () => {
                        if (selCountEl) selCountEl.textContent = xeroSelected.size;
                        if (confirmBtn) confirmBtn.disabled = xeroSelected.size === 0;
                        if (warningEl) warningEl.style.display = xeroSelected.size >= maxAllowed ? "inline" : "none";
                    };

                    if (listEl) {
                        listEl.innerHTML = "";
                        xeroConns.forEach(c => {
                            const isChecked = xeroSelected.has(c.companyId);
                            const row = document.createElement("div");
                            row.className = `xero-company-row ${isChecked ? "selected" : ""}`;
                            row.id = "xero_row_" + c.companyId;
                            row.innerHTML = `
                                    <input type="checkbox" class="xero-company-cb" id="xero_cb_${c.companyId}" value="${c.companyId}" ${isChecked ? "checked" : ""} />
                                    <div class="xero-company-icon">xero</div>
                                    <div class="xero-company-info">
                                        <div class="xero-company-name">${c.companyName || "Xero Organisation"}</div>
                                        <div class="xero-company-id">Tenant: ${c.companyId || "—"}</div>
                                    </div>
                                `;

                            const cb = row.querySelector(".xero-company-cb");

                            row.addEventListener("click", (e) => {
                                if (e.target === cb) return;
                                if (cb.checked) {
                                    cb.checked = false;
                                } else {
                                    if (xeroSelected.size >= maxAllowed && !xeroSelected.has(c.companyId)) return;
                                    cb.checked = true;
                                }
                                cb.dispatchEvent(new Event("change"));
                            });

                            cb.addEventListener("change", () => {
                                if (cb.checked) {
                                    if (xeroSelected.size >= maxAllowed) {
                                        cb.checked = false;
                                        return;
                                    }
                                    xeroSelected.add(c.companyId);
                                    row.classList.add("selected");
                                } else {
                                    xeroSelected.delete(c.companyId);
                                    row.classList.remove("selected");
                                }
                                updateXeroUI();
                            });

                            listEl.appendChild(row);
                        });
                    }
                    updateXeroUI();

                    if (confirmBtn) {
                        const newConfirmBtn = confirmBtn.cloneNode(true);
                        confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);

                        newConfirmBtn.addEventListener("click", async () => {
                            if (xeroSelected.size === 0) return;
                            newConfirmBtn.disabled = true;
                            newConfirmBtn.innerHTML = '<span class="spinner" style="display:inline-block;width:12px;height:12px;border:2px solid rgba(255,255,255,0.4);border-top-color:#fff;border-radius:50%;animation:spin 0.7s linear infinite;margin-right:6px;vertical-align:middle;"></span> Saving...';

                            try {
                                for (const c of xeroConns) {
                                    if (xeroSelected.has(c.companyId)) {
                                        await ApiService.apiFetch(`/api/connections/${c.companyId}/activate`, { method: "POST" });
                                    } else {
                                        await ApiService.apiFetch(`/api/connections/${c.companyId}`, { method: "DELETE" });
                                    }
                                }

                                const activeId = Array.from(xeroSelected)[0];
                                if (activeId) AppState.currentCompanyId = activeId;

                                DashboardService.showStatus("Xero companies saved successfully.", "success", null, "xero");
                                DashboardService.renderERPSection();
                            } catch (err) {
                                console.error("Error saving Xero companies:", err);
                                DashboardService.showStatus("Failed to save Xero companies.", "error", null, "xero");
                                newConfirmBtn.disabled = false;
                                newConfirmBtn.textContent = "Connect Selected Companies";
                            }
                        });
                    }
                } else {
                    if (selectXeroCard) selectXeroCard.style.display = "none";
                    if (activeCompanyCard) activeCompanyCard.style.display = "block";
                }

                // Populate Active Company Dropdown (platform-filtered)
                if (dropdown) {
                    dropdown.innerHTML = "";
                    const platformDropdownConns = activeConns.filter(c => (c.platform || "quickbooks").toLowerCase() === (AppState.currentProvider || "quickbooks"));
                    platformDropdownConns.forEach(c => {
                        const opt = document.createElement("option");
                        opt.value = c.companyId;
                        opt.dataset.platform = (c.platform || "QuickBooks").toLowerCase();
                        opt.textContent = `${c.companyName || "Company"}`;
                        if (c.companyId === AppState.currentCompanyId) opt.selected = true;
                        dropdown.appendChild(opt);
                    });
                }

                // Set active company details
                const activeConn = activeConns.find(c => c.companyId === AppState.currentCompanyId) || activeConns[0];
                if (activeConn) {
                    AppState.currentCompanyId = activeConn.companyId;
                    AppState.currentProvider = (activeConn.platform || "quickbooks").toLowerCase();
                    AppState.erpType = AppState.currentProvider;

                    // Persist the resolved active company so a refresh
                    // (which re-creates AppState from scratch) restores
                    // the same company instead of always falling back
                    // to the first one in the list. This is the single
                    // point every switch path (radio button, dropdown)
                    // routes back through via renderERPSection(), so it
                    // covers both QuickBooks and Xero uniformly.
                    localStorage.setItem("fa_current_company_id", activeConn.companyId || "");

                    const activeRealmEl = document.getElementById("activeRealmId");
                    if (activeRealmEl) activeRealmEl.textContent = activeConn.companyId || "—";
                    const realmEl = document.getElementById("connRealmId");
                    if (realmEl) realmEl.textContent = activeConn.companyId || "—";
                }

                // Platform-scoped connections (only show same platform as current active company)
                const currentPlatform = AppState.currentProvider || "quickbooks";
                const platformConns = conns.filter(c => (c.platform || "quickbooks").toLowerCase() === currentPlatform);

                // Populate Company Management List (platform-filtered)
                if (companyListEl) {
                    companyListEl.innerHTML = "";
                    platformConns.forEach(c => {
                        const isActive = c.companyId === AppState.currentCompanyId;
                        const isXero = (c.platform || "").toLowerCase() === "xero";
                        const isDisconnected = c.status === 'Disconnected';
                        const item = document.createElement("div");
                        item.className = `fa-company-item ${isActive ? "active-company" : ""} ${isDisconnected ? "disconnected-company" : ""}`;
                        item.dataset.companyId = c.companyId;

                        let badgeOrBtn = "";
                        if (isActive && !isDisconnected) {
                            badgeOrBtn = '<span class="fa-badge-active">ACTIVE</span>';
                        } else if (isDisconnected) {
                            badgeOrBtn = '<button class="fa-btn-reconnect">Reconnect</button>';
                        } else {
                            badgeOrBtn = '<button class="fa-btn-switch">Switch</button>';
                        }

                        item.innerHTML = `
                                <input type="radio" name="companyRadio" class="fa-company-radio" ${isActive && !isDisconnected ? "checked" : ""} />
                                <div class="fa-company-icon ${isXero ? 'xero-company-icon' : ''}">${isXero ? 'xero' : 'qb'}</div>
                                <div class="fa-company-info">
                                    <div class="fa-company-name">${c.companyName || (isXero ? "Xero Organisation" : "QuickBooks Company")}</div>
                                    <div class="fa-company-tag">Last Sync: ${this.formatRelativeTime(c.lastSyncedAt, c.status)}</div>
                                </div>
                                <div class="fa-company-actions">
                                    ${badgeOrBtn}
                                    <button class="fa-btn-dots" title="More options">⋮</button>
                                </div>
                            `;

                        // Click radio or card row to make active (ignore for disconnected companies)
                        item.addEventListener("click", (e) => {
                            if (e.target.classList.contains("fa-btn-dots")) {
                                e.stopPropagation();
                                this.showContextMenu(e.target, c);
                                return;
                            }
                            if (e.target.classList.contains("fa-btn-reconnect")) {
                                e.stopPropagation();
                                const reconnectPlatform = (c.platform || "quickbooks").toLowerCase();
                                this.showStatus("Launching re-authorization...", "success", null, reconnectPlatform);
                                // Pass the company this button belongs to, so
                                // the backend can hold the flow to it: picking
                                // a different company in the provider's
                                // account chooser must fail rather than being
                                // accepted as a new connection.
                                this.launchERPOAuth(reconnectPlatform, c.companyId);
                                return;
                            }
                            if (isDisconnected) {
                                // Do not activate disconnected companies on box click
                                return;
                            }
                            this.switchActiveCompany(c.companyId, platformConns);
                        });

                        companyListEl.appendChild(item);
                    });
                }

                if (footerEl) {
                    const platformLabel = currentPlatform === "xero" ? "Xero" : "QuickBooks";
                    footerEl.textContent = `Showing ${platformConns.length} ${platformLabel} ${platformConns.length === 1 ? "company" : "companies"}`;
                }

                // Populate Subscription Stats (filtered by current active platform: quickbooks or xero)
                const currentPlan = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null') ? AppState.subscriptionPlan : "Basic";
                const maxAllowed = getMaxCompaniesForPlan(currentPlan);

                const connectedCount = platformConns.length;
                const remaining = Math.max(0, maxAllowed - connectedCount);

                if (document.getElementById("subInfoPlan")) document.getElementById("subInfoPlan").textContent = currentPlan;
                if (document.getElementById("subInfoConnected")) document.getElementById("subInfoConnected").textContent = `${connectedCount} / ${maxAllowed}`;
                if (document.getElementById("subInfoRemaining")) document.getElementById("subInfoRemaining").textContent = String(remaining);

                // Update Tier Badge in Header
                const tierBadge = document.getElementById("connTierBadge");
                if (tierBadge) tierBadge.textContent = `${currentPlan.toUpperCase()} PLAN`;

                // Update dynamic button label for Pull Master Data
                const isQB = AppState.currentProvider === "quickbooks";
                const platformDisplayName = isQB ? "QuickBooks" : "Xero";
                const pullLabel = document.getElementById("pullBtnLabel");
                if (pullLabel) pullLabel.textContent = isQB ? "QBO" : "Xero";

                // Update "Company Management" section title to reflect platform
                const sectionTitle = document.querySelector(".fa-section-title");
                if (sectionTitle) sectionTitle.textContent = `${platformDisplayName} Companies`;

                // Update header status realm label (kept beside the ID
                // itself, not as a separate label elsewhere in the row)
                const connStatus = document.querySelector(".fa-conn-status");
                if (connStatus && activeConn) {
                    const idLabel = isQB ? "Realm ID" : "Tenant ID";
                    connStatus.innerHTML = `<span class="fa-realm-label">${idLabel}:</span> <span id="connRealmId">${activeConn.companyId || "—"}</span>`;
                }

                // Update Disconnect button label
                const disconnectBtn = document.getElementById("btnDisconnectERP");
                if (disconnectBtn) disconnectBtn.textContent = `Disconnect ${platformDisplayName}`;

                // Show console for correct provider
                const qbConsole = document.getElementById("qbConsole");
                const xeroConsole = document.getElementById("xeroConsole");
                if (qbConsole) qbConsole.style.display = isQB ? "flex" : "none";
                if (xeroConsole) xeroConsole.style.display = isQB ? "none" : "flex";

                // The active provider/company may have just changed
                // (switch/resume/connect) — recompute the badge/drawer,
                // log console, and step indicators so all three stay
                // scoped to exactly what's active now, never a leftover
                // from before.
                NotificationService.refreshForContext();
                this.renderActiveLogConsole();
                this.applyStepState();
            })
            .catch(() => {
                // Fallback to offline/disconnected view
                if (discSection) {
                    discSection.style.display = "flex";
                    discSection.style.flexDirection = "column";
                    discSection.style.height = "100%";
                }
                if (provSection) provSection.style.display = "none";
                if (connSection) connSection.style.display = "none";
                if (connectingSection) connectingSection.style.display = "none";
            });
    },

    formatRelativeTime(dateInput, status) {
        if (status === 'Disconnected') return "Disconnected";
        // A freshly connected company starts with status 'Not Synced'
        // and no lastSyncedAt — show that literally instead of a vaguer
        // "not synced yet". Once the first Master Data Pull succeeds,
        // the backend flips status to 'Active' and stamps lastSyncedAt,
        // so this falls through to the relative-time formatting below
        // (which reports "Just now" immediately after that pull).
        if (status === 'Not Synced' || !dateInput) return "Not Synced";
        const date = new Date(dateInput);
        if (isNaN(date.getTime())) return "Not Synced";
        const now = new Date();
        const diffMs = now - date;
        if (diffMs < 0) return "Just now";
        const diffSec = Math.floor(diffMs / 1000);
        if (diffSec < 45) return "Just now";
        const diffMin = Math.floor(diffSec / 60);
        if (diffMin < 60) {
            return diffMin === 1 ? "1 minute ago" : `${diffMin} minutes ago`;
        }
        const diffHr = Math.floor(diffMin / 60);
        if (diffHr < 24) {
            return diffHr === 1 ? "1 hour ago" : `${diffHr} hours ago`;
        }
        const diffDays = Math.floor(diffHr / 24);
        if (diffDays < 30) {
            return diffDays === 1 ? "1 day ago" : `${diffDays} days ago`;
        }
        return date.toLocaleDateString();
    },

    switchActiveCompany(companyId, conns) {
        AppState.currentCompanyId = companyId;

        // Set the provider immediately from the target company so renderERPSection shows correct platform
        const targetConn = conns.find(c => c.companyId === companyId);
        const targetPlatform = targetConn ? (targetConn.platform || "quickbooks").toLowerCase() : null;
        if (targetConn) {
            AppState.currentProvider = targetPlatform;
            AppState.erpType = AppState.currentProvider;
        }

        ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));

        // Toast only fires once the backend confirms the switch — not
        // when the click merely starts the request.
        ApiService.apiFetch(`/api/connections/${companyId}/activate`, { method: "POST" })
            .then(res => res.json())
            .then(data => {
                this.renderERPSection();
                if (targetConn) {
                    const countMsg = data.totalRecords ? ` (Total Records Found: ${data.totalRecords})` : "";
                    this.addLog(`Switched active company to: ${targetConn.companyName}${countMsg}`);
                    this.showStatus(
                        `Active company updated to ${targetConn.companyName}`,
                        "success",
                        data.totalRecords !== undefined ? `Total Records Found: ${data.totalRecords}` : null,
                        targetPlatform
                    );
                }
            })
            .catch(err => {
                console.error("Error activating company:", err);
                this.renderERPSection();
                this.showStatus("Failed to switch active company.", "error", null, targetPlatform);
            });
    },

    showContextMenu(targetBtn, company) {
        const menu = document.getElementById("companyContextMenu");
        if (!menu) return;
        const rect = targetBtn.getBoundingClientRect();
        menu.style.top = `${rect.bottom + 4}px`;
        menu.style.left = `${rect.left - 100}px`;
        menu.style.display = "block";

        const editBtn = document.getElementById("ctxEdit");
        const disconnectBtn = document.getElementById("ctxDisconnect");

        if (editBtn) {
            editBtn.onclick = () => {
                menu.style.display = "none";
                this.showRenameModal(company);
            };
        }
        if (disconnectBtn) {
            disconnectBtn.onclick = async () => {
                menu.style.display = "none";
                const companyPlatform = (company.platform || "quickbooks").toLowerCase();
                this.showStatus(`Disconnecting ${company.companyName}...`, "success", null, companyPlatform);
                try {
                    await ApiService.apiFetch(`/api/connections/${company.companyId}`, { method: "DELETE" });
                    // If we disconnected the currently active company, reset it so a new one is picked
                    if (AppState.currentCompanyId === company.companyId) {
                        AppState.currentCompanyId = null;
                    }
                    try { await ExcelService.clearMasterData(); } catch (_) { }
                    this.showStatus("Company disconnected.", "success", null, companyPlatform);
                    this.renderERPSection();
                } catch (_) {
                    this.showStatus("Failed to disconnect company.", "error", null, companyPlatform);
                }
            };
        }
    },

    showRenameModal(company) {
        const modal = document.getElementById("renameCompanyModal");
        const input = document.getElementById("renameCompanyInput");
        const closeBtn = document.getElementById("btnCloseRenameCompany");
        const cancelBtn = document.getElementById("btnCancelRenameCompany");
        const confirmBtn = document.getElementById("btnConfirmRenameCompany");

        if (!modal || !input) return;
        input.value = company.companyName || "";
        modal.style.display = "flex";
        input.focus();

        const closeModal = () => {
            modal.style.display = "none";
        };

        if (closeBtn) closeBtn.onclick = closeModal;
        if (cancelBtn) cancelBtn.onclick = closeModal;

        if (confirmBtn) {
            confirmBtn.onclick = async () => {
                const newName = input.value.trim();
                if (!newName) return;
                const companyPlatform = (company.platform || "quickbooks").toLowerCase();
                confirmBtn.disabled = true;
                try {
                    const res = await ApiService.apiFetch(`/api/connections/${company.companyId}/rename`, {
                        method: "PATCH",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ companyName: newName })
                    });
                    if (res.ok) {
                        company.companyName = newName;
                        this.showStatus("Company renamed successfully.", "success", null, companyPlatform);
                        closeModal();
                        this.renderERPSection();
                    } else {
                        this.showStatus("Failed to rename company.", "error", null, companyPlatform);
                    }
                } catch (err) {
                    console.error("Rename error:", err);
                    this.showStatus("Failed to rename company.", "error", null, companyPlatform);
                } finally {
                    confirmBtn.disabled = false;
                }
            };
        }
    },

    ...createDashboardConsole(),
    ...createErpConnection()
};

export { DashboardService };
