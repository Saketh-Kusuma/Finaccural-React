/**
 * Dashboard data actions: Setup Sheets, Pull Master Data, Refresh Schedule,
 * the journal-type toggles and the provider/tier tab dropdown.
 *
 * Split out of bindDashboardView() unchanged and still invoked from it, so
 * the listeners are registered in exactly the original order.
 *
 * Pull Master Data and Refresh Schedule are the two halves of one
 * server-driven pagination cursor: Pull always restarts a cycle, Refresh
 * always continues it. See the comments on each handler.
 */
import { AppState } from "../state/appState.js";
import { ApiService } from "../services/apiService.js";
import { ExcelService } from "../services/excelService.js";
import { DashboardService } from "../services/dashboardService.js";
import { AppController } from "./appController.js";
import { flattenAllMasterDataRecords } from "../services/excelDataMappers.js";
import {
    getPullPageCursor,
    setPullPageCursor,
    clearPullPageCursor
} from "../batchDataLoader.js";
import { ERROR_CODES } from "../../shared/apiErrorHandler.js";

function checkTrialExpiredGuard(provider) {
    if (AppController.isTrialExpired()) {
        const msg = "Your free trial has expired. Please upgrade your plan to continue.";
        DashboardService.addLog(`Action failed: ${msg}`);
        DashboardService.showStatus("Trial Expired", "error", "Please upgrade your plan to perform data actions.", provider);
        const modal = document.getElementById("trialExpiredModal");
        if (modal) modal.style.display = "flex";
        return true;
    }
    return false;
}

export function bindDataActionHandlers() {

    // Setup Sheets button in provider-selected state
    document.getElementById("setupBtnProv")?.addEventListener("click", async () => {
        if (checkTrialExpiredGuard(AppState.currentProvider)) return;
        try {
            document.getElementById("provStepSetup")?.classList.add("active");
            DashboardService.addLog(`Setting up Master & Input sheets for ${AppState.currentProvider === "quickbooks" ? "QuickBooks" : "Xero"}...`);
            DashboardService.showStatus("Setting up sheets...", "success", null, AppState.currentProvider);
            await ExcelService.setupWorkbookSheets(AppState.currentProvider);
            DashboardService.markStepComplete("setup");
            DashboardService.addLog("Sheets setup successfully.");
            DashboardService.showStatus("Master and Input sheets setup successfully.", "success", null, AppState.currentProvider);
        } catch (error) {
            console.error(error);
            DashboardService.addLog("Error setting up sheets: " + error.message);
            DashboardService.showStatus("Error setting up sheets.", "error", null, AppState.currentProvider);
        }
    });

    // Pull Data button in provider-selected state
    // Pull Master Data buttons (provider-selected + default state
    // share this one handler, same as Refresh Schedule further
    // below).
    //
    // Pull Master Data ALWAYS STARTS OVER. It is the "begin a new
    // pull" button, not a "continue" button: every click discards
    // the stored pagination cursor, wipes the sheet's master-data
    // range, and asks the backend for the first 100 records of the
    // first API — no matter how far a previous cycle had already
    // progressed. Clicking it halfway through a cycle is therefore
    // indistinguishable from clicking it for the very first time,
    // which is what guarantees no duplicated and no skipped rows:
    // the sheet and the cursor are reset together, in the same
    // click, so neither can outlive the other.
    //
    // Refresh Schedule (further below) is the CONTINUE button: it
    // reads that same stored cursor and asks for the NEXT 100
    // records, appending them.
    //
    // Either way a click fetches exactly one batch of 100 records
    // of ONE entity. The backend walks the entities strictly one
    // at a time — Accounts first, 10 at a time until Accounts is
    // completely finished, then Classes from its own first record,
    // then Locations, then Customers, then Vendors — so no two
    // APIs are ever fetched at the same time. The backend decides
    // what "one batch" contains via a real MAXRESULTS=100
    // QuickBooks request for that single entity; this handler
    // never fetches everything and slices it client-side.
    const handlePullClick = async (event) => {
        const button = event.currentTarget;
        const isProv = button.id === "pullBtnProv";
        const provider = AppState.currentProvider;
        const companyId = AppState.currentCompanyId;
        const providerLabel = provider === "quickbooks" ? "QuickBooks" : "Xero";

        if (checkTrialExpiredGuard(provider)) return;

        // Prerequisite check: user MUST complete Setup Master & Input Sheets before pulling master data
        if (!DashboardService.isStepComplete("setup")) {
            const detailMsg = `Cannot pull master data: You must run Setup Master & Input Sheets for ${providerLabel} first.`;
            DashboardService.addLog(`Pull Master Data failed: ${detailMsg}`);
            DashboardService.showStatus("Pull Master Data Failed", "error", detailMsg, provider);
            return;
        }

        const stepId = isProv
            ? "provStepPull"
            : (provider === "quickbooks" ? "stepPull" : "xeroStepPull");

        // Prevent duplicate clicks
        if (button.disabled) return;
        button.disabled = true;

        try {
            document.getElementById(stepId)?.classList.add("active");

            DashboardService.addLog(`Pulling master data from ${providerLabel}...`);
            DashboardService.showStatus("Initializing Data Pull...", "success", "Pre-flight record count check in progress...", provider);

            const hadCursor = !!getPullPageCursor(provider, companyId);
            clearPullPageCursor(provider, companyId);
            if (hadCursor) {
                DashboardService.addLog("Pull Master Data: restarting from the first batch — clearing previously pulled data.");
            }

            const isFreshCycle = true;

            const onProgress = (event) => {
                if (event.type === 'start') {
                    DashboardService.showStatus(
                        `Pulling Master Data (0%)`,
                        "success",
                        `Total Records Found: ${event.totalRecords}. Starting fetch...`,
                        provider
                    );
                } else if (event.type === 'progress') {
                    const pct = event.percentage || 0;
                    DashboardService.showStatus(
                        `Pulling Master Data (${pct}%)`,
                        "success",
                        `Fetched ${event.fetchedRecords} of ${event.totalRecords} records`,
                        provider
                    );
                }
            };

            const data = await ApiService.fetchMasterDataStream(provider, companyId, onProgress);

            await ExcelService.clearMasterDataRange();
            DashboardService.markStepIncomplete("pull");

            const batch = flattenAllMasterDataRecords(data, { includeCompany: isFreshCycle });

            if (batch.length === 0 && isFreshCycle) {
                clearPullPageCursor(provider, companyId);
                DashboardService.markStepComplete("pull");
                DashboardService.addLog("Pull: no more data available.");
                DashboardService.showStatus("No more data available.", "success", "No master data found for this company.", provider);
                DashboardService.renderERPSection();
                return;
            }

            await ExcelService.appendManualBatch(provider, batch);

            clearPullPageCursor(provider, companyId);
            DashboardService.markStepComplete("pull");

            const counts = {
                company: Array.isArray(data.company) ? data.company.length : (data.company ? 1 : 0),
                accounts: Array.isArray(data.accounts) ? data.accounts.length : 0,
                classes: Array.isArray(data.classes) ? data.classes.length : 0,
                locations: Array.isArray(data.locations) ? data.locations.length : 0,
                customers: Array.isArray(data.customers) ? data.customers.length : 0,
                vendors: Array.isArray(data.vendors) ? data.vendors.length : 0
            };
            const totalRecords = counts.accounts + counts.classes + counts.locations + counts.customers + counts.vendors;

            const pullTitle = "Data completed.";
            const pullDetail = `Successfully fetched all ${totalRecords} records across all entities for this company.`;
            DashboardService.addLog(`${pullTitle} (${totalRecords} records pulled)`);
            DashboardService.showStatus(pullTitle, "success", pullDetail, provider);
            DashboardService.renderERPSection();
        } catch (error) {
            console.error(error);
            const isExpired = error.code === ERROR_CODES.ERP_SESSION_EXPIRED;
            const msg = isExpired
                ? error.message
                : "Error pulling data: " + error.message;
            DashboardService.addLog(msg);
            DashboardService.showStatus(
                isExpired ? error.message : (isProv ? "Data pull failed." : "Please set up the master sheet before pulling the master data"),
                "error",
                isExpired ? "" : (isProv ? "Please try again." : ""),
                provider
            );
            if (isExpired) {
                try {
                    DashboardService.renderERPSection();
                } catch (renderErr) {
                    console.error("Failed to refresh ERP section after session expiry:", renderErr);
                }
            }
        } finally {
            button.disabled = false;
        }
    };

    document.getElementById("pullBtnProv")?.addEventListener("click", handlePullClick);

    // Journal type buttons in provider-selected state
    document.querySelectorAll("#dashProviderSelected .conn-journal-btn").forEach(btn => {
        btn.addEventListener("click", () => {
            document.querySelectorAll("#dashProviderSelected .conn-journal-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
        });
    });

    // Setup Sheets button
    document.getElementById("setupBtn")?.addEventListener("click", async () => {
        if (checkTrialExpiredGuard(AppState.currentProvider)) return;
        try {
            const stepSetupId = AppState.currentProvider === "quickbooks" ? "stepSetup" : "xeroStepSetup";
            document.getElementById(stepSetupId)?.classList.add("active");

            DashboardService.addLog(`Setting up Master & Input sheets for ${AppState.currentProvider === "quickbooks" ? "QuickBooks" : "Xero"}...`);
            DashboardService.showStatus("Setting up sheets...", "success", null, AppState.currentProvider);
            await ExcelService.setupWorkbookSheets(AppState.currentProvider);
            DashboardService.completeStep("stepSetup");
            DashboardService.addLog("Sheets setup successfully.");
            DashboardService.showStatus("Master and Input sheets setup successfully.", "success", null, AppState.currentProvider);
        } catch (error) {
            console.error(error);
            DashboardService.addLog("Error setting up sheets: " + error.message);
            DashboardService.showStatus("Error setting up sheets.", "error", null, AppState.currentProvider);
        }
    });

    // Pull Data button
    document.getElementById("pullBtn")?.addEventListener("click", handlePullClick);

    // Journal type buttons
    document.querySelectorAll(".conn-journal-btn").forEach(btn => {
        btn.addEventListener("click", () => {
            document.querySelectorAll(".conn-journal-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
        });
    });

    // Provider Tab Toggles
    const tabQB = document.getElementById("tabQB");
    const tabXero = document.getElementById("tabXero");
    const options = document.getElementById("connTierOptions");

    const toggleOptions = (provider) => {
        if (options) {
            options.style.display = options.style.display === "flex" ? "none" : "flex";
            options.style.flexDirection = "column";

            // Render correct options list
            const isQB = provider === "quickbooks";
            document.getElementById("tierBasic").textContent = isQB ? "Q Basic" : "Xero Basic";
            document.getElementById("tierStandard").textContent = isQB ? "Q Standard" : "Xero Standard";
            document.getElementById("tierPro").textContent = isQB ? "Q Pro" : "Xero Pro";
        }
    };

    tabQB?.addEventListener("click", (e) => {
        e.stopPropagation();
        if (AppState.erpConnected && AppState.erpType !== "quickbooks") return; // Tab locked to connection
        toggleOptions("quickbooks");
    });

    tabXero?.addEventListener("click", (e) => {
        e.stopPropagation();
        if (AppState.erpConnected && AppState.erpType !== "xero") return; // Tab locked to connection
        toggleOptions("xero");
    });

    // Dropdown option clicks
    document.querySelectorAll(".conn-tier-opt").forEach(opt => {
        opt.addEventListener("click", (e) => {
            e.stopPropagation();
            document.querySelectorAll(".conn-tier-opt").forEach(o => o.classList.remove("selected"));
            opt.classList.add("selected");

            const tierName = opt.textContent;
            AppState.erpTier = tierName;
            const badge = document.getElementById("connTierBadge");
            if (badge) badge.textContent = tierName;

            if (options) options.style.display = "none";
        });
    });

    // Collapse dropdown on outside click
    document.addEventListener("click", () => {
        if (options) options.style.display = "none";
    });

    // Refresh Schedule buttons
    const handleRefreshClick = async (event) => {
        const button = event.currentTarget;

        const provider = AppState.currentProvider;
        const companyId = AppState.currentCompanyId;
        const providerLabel = provider === "quickbooks" ? "QuickBooks" : "Xero";

        if (checkTrialExpiredGuard(provider)) return;

        // Prerequisite check: user MUST complete Setup Master & Input Sheets and Pull Master Data first
        const isSetupDone = DashboardService.isStepComplete("setup");
        const isPullDone = DashboardService.isStepComplete("pull");

        if (!isSetupDone || !isPullDone) {
            const missingSteps = [];
            if (!isSetupDone) missingSteps.push("Setup Master & Input Sheets");
            if (!isPullDone) missingSteps.push("Pull Master Data");
            const detailMsg = `Cannot refresh schedule: You must run ${missingSteps.join(" and ")} from ${providerLabel} first.`;

            DashboardService.addLog(`Refresh Schedule failed: ${detailMsg}`);
            DashboardService.showStatus("Refresh Schedule Failed", "error", detailMsg, provider);
            return;
        }

        const icon = button.querySelector(".refresh-icon");
        if (icon) icon.classList.add("spin");

        try {
            DashboardService.addLog(`Refreshing live data from ${providerLabel}...`);
            DashboardService.showStatus("Refreshing...", "success", null, provider);

            // Fetch complete master data stream to compare against existing sheet records
            const data = await ApiService.fetchMasterDataStream(provider, companyId);
            const batch = flattenAllMasterDataRecords(data, { includeCompany: false });

            // Append batch without duplicates and return exact new record count written
            const updatedCount = await ExcelService.appendManualBatch(provider, batch);

            const timestamp = new Date().toLocaleTimeString();
            await ExcelService.stampLastRefreshed(timestamp);

            clearPullPageCursor(provider, companyId);

            const refreshTitle = "Schedule Refreshed";
            if (updatedCount === 0) {
                DashboardService.addLog("Refresh Schedule complete: 0 updated records added.");
                DashboardService.showStatus(refreshTitle, "success", "0 updated records added.", provider);
            } else {
                const refreshDetail = `${updatedCount} updated record${updatedCount === 1 ? "" : "s"} added.`;
                DashboardService.addLog(`Refresh Schedule complete: ${refreshDetail}`);
                DashboardService.showStatus(refreshTitle, "success", refreshDetail, provider);
            }
        } catch (err) {
            console.error("Refresh error:", err);
            DashboardService.addLog("Error refreshing: " + err.message);
            DashboardService.showStatus("Data refresh failed.", "error", "Please try again.", provider);
        } finally {
            setTimeout(() => {
                if (icon) icon.classList.remove("spin");
            }, 1000);
        }
    };

    document.getElementById("btnRefreshScheduleProv")?.addEventListener("click", handleRefreshClick);
    document.getElementById("btnRefreshSchedule")?.addEventListener("click", handleRefreshClick);
}
