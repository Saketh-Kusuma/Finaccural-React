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
import { flattenAllMasterDataRecords } from "../services/excelDataMappers.js";
import {
    getPullPageCursor,
    setPullPageCursor,
    clearPullPageCursor
} from "../batchDataLoader.js";
import { ERROR_CODES } from "../../shared/apiErrorHandler.js";

export function bindDataActionHandlers() {

    // Setup Sheets button in provider-selected state
    document.getElementById("setupBtnProv")?.addEventListener("click", async () => {
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
    // range, and asks the backend for the first 10 records of the
    // first API — no matter how far a previous cycle had already
    // progressed. Clicking it halfway through a cycle is therefore
    // indistinguishable from clicking it for the very first time,
    // which is what guarantees no duplicated and no skipped rows:
    // the sheet and the cursor are reset together, in the same
    // click, so neither can outlive the other.
    //
    // Refresh Schedule (further below) is the CONTINUE button: it
    // reads that same stored cursor and asks for the NEXT 10
    // records, appending them.
    //
    // Either way a click fetches exactly one batch of 10 records
    // of ONE entity. The backend walks the entities strictly one
    // at a time — Accounts first, 10 at a time until Accounts is
    // completely finished, then Classes from its own first record,
    // then Locations, then Customers, then Vendors — so no two
    // APIs are ever fetched at the same time. The backend decides
    // what "one batch" contains via a real MAXRESULTS=10
    // QuickBooks request for that single entity; this handler
    // never fetches everything and slices it client-side.
    const handlePullClick = async (event) => {
        const button = event.currentTarget;
        const isProv = button.id === "pullBtnProv";
        const provider = AppState.currentProvider;
        const companyId = AppState.currentCompanyId;
        const providerLabel = provider === "quickbooks" ? "QuickBooks" : "Xero";
        const stepId = isProv
            ? "provStepPull"
            : (provider === "quickbooks" ? "stepPull" : "xeroStepPull");

        try {
            document.getElementById(stepId)?.classList.add("active");

            DashboardService.addLog(`Pulling master data from ${providerLabel}...`);
            DashboardService.showStatus("Pulling data...", "success", null, provider);

            // Pull Master Data is unconditionally a fresh start.
            // Drop any cursor left behind by an in-progress cycle
            // BEFORE the request goes out, so that even if the
            // fetch below fails the next click still begins at
            // record 1 rather than resuming a stale position.
            // Passing null as the cursor makes the backend reset
            // every API's offset/page/cursor to the beginning.
            const hadCursor = !!getPullPageCursor(provider, companyId);
            clearPullPageCursor(provider, companyId);
            if (hadCursor) {
                DashboardService.addLog("Pull Master Data: restarting from the first batch — clearing previously pulled data.");
            }

            // Every Pull click writes the cycle's first batch, so
            // the org header row is always seeded and the sheet is
            // always cleared first.
            const isFreshCycle = true;

            const data = await ApiService.fetchMasterData(provider, companyId, null);

            // Remove the previously pulled master data. Done after
            // the fetch succeeds, so a failed request never leaves
            // the user with an emptied sheet and nothing to show
            // for it.
            await ExcelService.clearMasterDataRange();

            // The previous cycle's "pull complete" tick no longer
            // describes what's on the sheet — this click has taken
            // it back to just the first batch.
            DashboardService.markStepIncomplete("pull");

            // CompanyInfo isn't paginated — the backend refetches
            // it on every click of a cycle, not just the first —
            // so only the cycle's first click seeds the org header
            // row; skip it on every later page to avoid a
            // duplicate, contentless "OrgName" row per click.
            const batch = flattenAllMasterDataRecords(data, { includeCompany: isFreshCycle });

            if (batch.length === 0 && isFreshCycle) {
                clearPullPageCursor(provider, companyId);
                DashboardService.markStepComplete("pull");
                DashboardService.addLog("Pull: no more data available.");
                DashboardService.showStatus("No more data available.", "success", "No master data found for this company.", provider);
                DashboardService.renderERPSection();
                return;
            }

            // One click = one batch, always — no exceptions for a
            // repeat cycle. Exactly ONE /api/pull-master-data
            // request was made above (ONE QuickBooks request
            // inside it, for the single entity currently being
            // drained, max 10 records); write just that response
            // and stop. The next batch — whether it's the same
            // entity's next 10 records or the first 10 of the next
            // entity in the order — is only fetched on the NEXT
            // click, never automatically within this one.
            await ExcelService.appendManualBatch(provider, batch);

            // A response that reports "not done" but carries no
            // cursor cannot be resumed — storing it would leave
            // every later click restarting the cycle at the first
            // entity's first record while forever reporting
            // "Batch written.". Treat that as the end of the cycle
            // instead, so the flow can never livelock.
            const pullFinished = data.isDone || !data.cursor;

            if (pullFinished) {
                clearPullPageCursor(provider, companyId);
                DashboardService.markStepComplete("pull");
            } else {
                setPullPageCursor(provider, companyId, data.cursor);
            }

            const pullTitle = pullFinished ? "Data completed." : "Batch written.";
            // No row-range numbers (e.g. "Rows 71-80 of 150") in the
            // user-facing detail — just the plain outcome/next step.
            // Finished state is just "Data completed." on its own,
            // no extra detail line.
            const pullDetail = pullFinished ? "" : "Click Pull Master Data again for the next batch.";
            DashboardService.addLog(pullDetail ? `${pullTitle} ${pullDetail}` : pullTitle);
            DashboardService.showStatus(pullTitle, "success", pullDetail || null, provider);
            DashboardService.renderERPSection();
        } catch (error) {
            console.error(error);
            // ApiService already showed the orange "reconnect" banner
            // (or the offline banner) as a global side effect when
            // this came from apiFetch — branch on the standardized
            // `.code` here too, never on message text.
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
                // renderERPConsole() only toggles a progress-step
                // marker — it doesn't touch the company badge. To
                // actually flip ACTIVE -> Reconnect in the UI, we
                // need to re-fetch connections and re-render the
                // company list, which is what renderERPSection()
                // does. Guarded in case of an unexpected error, so
                // it can't surface as an uncaught runtime popup.
                try {
                    DashboardService.renderERPSection();
                } catch (renderErr) {
                    console.error("Failed to refresh ERP section after session expiry:", renderErr);
                }
            }
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

        // Refresh Schedule no longer requires Setup/Pull to already
        // be marked "complete" before it can run — Pull Master Data
        // and Refresh are two triggers for the same server-driven
        // pagination cursor (see handlePullClick above and
        // batchDataLoader.js#getPullPageCursor), so Refresh must be
        // usable right after the very first Pull click, not just
        // once an entire multi-page pull cycle has fully drained.

        const icon = button.querySelector(".refresh-icon");
        if (icon) icon.classList.add("spin");

        const provider = AppState.currentProvider;
        const companyId = AppState.currentCompanyId;
        const providerLabel = provider === "quickbooks" ? "QuickBooks" : "Xero";

        try {
            DashboardService.addLog(`Refreshing live data from ${providerLabel}...`);
            DashboardService.showStatus("Refreshing...", "success", null, provider);

            // Refresh Schedule is the CONTINUE half of the pair:
            // it reads the same stored per-provider/company cursor
            // Pull Master Data writes (see handlePullClick above)
            // and asks for the NEXT batch of 10 records of the ONE
            // entity currently being drained, appending it to
            // what's already on the sheet — 1-10, then 11-20, then
            // 21-30, and so on through the fixed Accounts ->
            // Classes -> Locations -> Customers -> Vendors order.
            //
            // Unlike Pull Master Data, Refresh never resets the
            // cursor. The one case where it does start over is
            // when there is no cursor at all — nothing has been
            // pulled yet, or the last cycle already finished — in
            // which case there is no position to continue from and
            // a new cycle begins at record 1, clearing the sheet
            // exactly as a Pull would.
            const priorCursor = getPullPageCursor(provider, companyId);
            const isFreshCycle = !priorCursor;

            const data = await ApiService.fetchMasterData(provider, companyId, priorCursor);

            if (isFreshCycle) {
                // Start of a fresh cycle: clear the sheet's data
                // range exactly once, right before writing this
                // click's first page, same as Pull Master Data.
                await ExcelService.clearMasterDataRange();
            }

            // CompanyInfo isn't paginated — refetched on every
            // click of a cycle, not just the first — so only the
            // cycle's first click seeds the org header row.
            const batch = flattenAllMasterDataRecords(data, { includeCompany: isFreshCycle });

            if (batch.length === 0 && isFreshCycle) {
                clearPullPageCursor(provider, companyId);
                const timestamp = new Date().toLocaleTimeString();
                await ExcelService.stampLastRefreshed(timestamp);
                DashboardService.addLog("Refresh: no more data available.");
                DashboardService.showStatus("No more data available.", "success", "No master data found for this company.", provider);
                return;
            }

            // One click = one batch, always — no exceptions for a
            // repeat cycle. Exactly ONE /api/pull-master-data
            // request was made above (ONE QuickBooks request
            // inside it, for the single entity currently being
            // drained, max 10 records); write just that response
            // and stop. The next batch — whether it's the same
            // entity's next 10 records or the first 10 of the next
            // entity in the order — is only fetched on the NEXT
            // click, never automatically within this one.
            await ExcelService.appendManualBatch(provider, batch);

            const timestamp = new Date().toLocaleTimeString();
            await ExcelService.stampLastRefreshed(timestamp);

            // Same no-cursor guard as handlePullClick above — see
            // the comment there for why a cursor-less "not done"
            // response has to end the cycle.
            const refreshFinished = data.isDone || !data.cursor;

            if (refreshFinished) {
                clearPullPageCursor(provider, companyId);
                DashboardService.markStepComplete("pull");
            } else {
                setPullPageCursor(provider, companyId, data.cursor);
            }

            const refreshTitle = refreshFinished ? "Data completed." : "Batch written.";
            // No row-range numbers (e.g. "Rows 71-80 of 150") in the
            // user-facing detail — just the plain outcome/next step.
            // Finished state is just "Data completed." on its own,
            // no extra detail line.
            const refreshDetail = refreshFinished ? "" : "Click Refresh again for the next batch.";
            DashboardService.addLog(refreshDetail ? `${refreshTitle} ${refreshDetail}` : refreshTitle);
            DashboardService.showStatus(refreshTitle, "success", refreshDetail || null, provider);
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
