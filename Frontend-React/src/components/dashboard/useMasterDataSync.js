import { useState, useCallback, useEffect, useRef } from "react";
import {
  fetchMasterDataStream,
  fetchIncrementalDataStream,
} from "../../taskpane/api";
import { ExcelService } from "../../taskpane/services/excelService";
import { flattenAllMasterDataRecords } from "../../taskpane/services/excelDataMappers";

export function useMasterDataSync({
  provider,
  activeConnection,
  realmId,
  planClean,
  label,
  notify,
  onTrialExpired,
}) {
  const [spinning, setSpinning] = useState(false);
  const [setupBusy, setSetupBusy] = useState(false);
  const [pullBusy, setPullBusy] = useState(false);

  const companyId = activeConnection?.companyId || "";

  // Per-company localStorage keys
  const setupKey = companyId ? `fa_step_setup_${companyId}` : "fa_step_setup";
  const pullKey  = companyId ? `fa_step_pull_${companyId}`  : "fa_step_pull";

  const [isSetupDone, setIsSetupDone] = useState(
    () => localStorage.getItem(companyId ? `fa_step_setup_${companyId}` : "fa_step_setup") === "complete"
  );
  const [isPullDone, setIsPullDone] = useState(
    () => localStorage.getItem(companyId ? `fa_step_pull_${companyId}` : "fa_step_pull") === "complete"
  );
  const [logs, setLogs] = useState([]);

  // When the active company changes, Excel data is cleared, so reset setup & pull steps to blue (not completed)
  const prevCompanyIdRef = useRef(companyId);
  useEffect(() => {
    if (prevCompanyIdRef.current === companyId) return;
    prevCompanyIdRef.current = companyId;

    if (companyId) {
      localStorage.removeItem(`fa_step_setup_${companyId}`);
      localStorage.removeItem(`fa_step_pull_${companyId}`);
    }
    localStorage.removeItem("fa_step_setup");
    localStorage.removeItem("fa_step_pull");

    setIsSetupDone(false);
    setIsPullDone(false);
    setLogs([]);
  }, [companyId]);

  const addLog = useCallback((msg) => {
    const time = new Date().toLocaleTimeString();
    setLogs((prev) => [...prev, `[${time}] ${msg}`]);
  }, []);

  const checkTrialExpiredGuard = () => {
    const currentPlan = (localStorage.getItem("fa_plan") || localStorage.getItem("fa_subscription_plan") || "").toLowerCase();
    const isPaid = currentPlan.includes("basic") || currentPlan.includes("standard") || currentPlan.includes("pro") || currentPlan.includes("enterprise");
    if (isPaid) return false;

    let isExpired = currentPlan === "expired";
    const trialEndsStr = localStorage.getItem("fa_trial_ends_at");
    if (trialEndsStr) {
      const num = Number(trialEndsStr);
      if (!isNaN(num) && num > 0 && Date.now() >= num) isExpired = true;
      const dateNum = new Date(trialEndsStr).getTime();
      if (!isNaN(dateNum) && dateNum > 0 && Date.now() >= dateNum) isExpired = true;
    }
    const trialStartStr = localStorage.getItem("fa_trial_start");
    if (trialStartStr) {
      const num = Number(trialStartStr);
      if (!isNaN(num) && num > 0 && Date.now() >= (num + 2 * 60 * 1000)) isExpired = true;
    }

    if (isExpired) {
      const msg = "Your free trial has expired. Please upgrade your plan to continue.";
      addLog(`Action failed: ${msg}`);
      notify("Trial Expired", "error", "Please upgrade your plan to perform data actions.", provider);
      ExcelService.clearMasterData().catch((e) => console.error("Failed to clear master data on trial expiry", e));
      if (onTrialExpired) onTrialExpired();
      return true;
    }
    return false;
  };

  const checkExpiredCompanyGuard = () => {
    if (activeConnection?.status === "Disconnected") {
      const compName = activeConnection?.companyName || "company";
      addLog(`Action blocked: Connection for ${compName} has expired. Please reconnect.`);
      notify(
        "Connection Expired",
        "error",
        `The connection for ${compName} has expired. Please click Reconnect in the company list to restore access.`,
        provider
      );
      return true;
    }
    return false;
  };

  const handleSetup = async () => {
    if (checkTrialExpiredGuard()) return;
    if (checkExpiredCompanyGuard()) return;
    if (setupBusy) return;
    setSetupBusy(true);
    addLog(`Setting up Master & Input sheets for ${label}...`);
    notify("Setting up sheets...", "success", null, provider);
    try {
      await ExcelService.setupWorkbookSheets(provider);
      localStorage.setItem(setupKey, "complete");
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

  const handlePull = async () => {
    if (checkTrialExpiredGuard()) return;
    if (checkExpiredCompanyGuard()) return;
    if (pullBusy) return;

    if (!isSetupDone) {
      const detailMsg = `Cannot pull master data: You must run Setup Master & Input Sheets for ${label} first.`;
      addLog(`Pull Master Data failed: ${detailMsg}`);
      notify("Pull Master Data Failed", "error", detailMsg, provider);
      return;
    }

    setPullBusy(true);
    const activeId = activeConnection?.companyId || realmId || "";

    // Start free trial timer when Pull Master Data is clicked if not already active
    const currentPlan = (localStorage.getItem("fa_plan") || localStorage.getItem("fa_subscription_plan") || "").toLowerCase();
    if (currentPlan.includes("trial") && !localStorage.getItem("fa_trial_start") && !localStorage.getItem("fa_trial_ends_at")) {
      const now = Date.now();
      localStorage.setItem("fa_trial_start", now.toString());
      localStorage.setItem("fa_trial_ends_at", (now + 2 * 60 * 1000).toString());
    }

    addLog(`Pulling master data from ${label}...`);
    notify("Initializing Data Pull...", "success", "Pre-flight record count check in progress...", provider);

    try {
      const data = await fetchMasterDataStream(
        provider,
        activeId,
        planClean,
        (progress) => {
          if (progress.percentage) {
            notify(
              `Pulling Master Data (${progress.percentage}%)`,
              "success",
              `Fetched ${progress.fetchedRecords || 0} of ${progress.totalRecords || 0} records`,
              provider
            );
          }
        }
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
        localStorage.setItem(pullKey, "complete");
        setIsPullDone(true);
        return;
      }

      const count = await ExcelService.appendManualBatch(provider, batch);
      localStorage.setItem(pullKey, "complete");
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

  const handleRefresh = async () => {
    if (checkTrialExpiredGuard()) return;
    if (checkExpiredCompanyGuard()) return;
    if (spinning) return;

    if (!isSetupDone) {
      const msg = `Cannot refresh: You must run Setup Master & Input Sheets for ${label} first.`;
      addLog(`Refresh failed: ${msg}`);
      notify("Refresh Failed", "error", msg, provider);
      return;
    }

    if (!isPullDone) {
      const msg = `Cannot refresh: You must Pull Master Data for ${label} before refreshing.`;
      addLog(`Refresh failed: ${msg}`);
      notify("Refresh Failed", "error", msg, provider);
      return;
    }

    setSpinning(true);
    const activeId = activeConnection?.companyId || realmId || "";
    addLog(`Refreshing live data from ${label}...`);
    notify("Refreshing...", "success", null, provider);
    try {
      const data = await fetchIncrementalDataStream(
        provider,
        activeId,
        planClean
      );
      if (data) {
        const batch = flattenAllMasterDataRecords(data, {
          includeCompany: false,
        });
        const updatedCount = await ExcelService.appendManualBatch(
          provider,
          batch
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

  return {
    spinning,
    setupBusy,
    pullBusy,
    isSetupDone,
    isPullDone,
    logs,
    addLog,
    handleSetup,
    handlePull,
    handleRefresh,
  };
}
