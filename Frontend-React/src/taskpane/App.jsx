import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Dashboard } from "../components/Dashboard";
import { Payment } from "../components/Payment";
import { Plans } from "../components/Plans";
import { TrialSelect } from "../components/TrialSelect";
import { TrialSelectModal } from "../components/TrialSelectModal";
import { Welcome } from "../components/Welcome";
import { Loading, Success } from "../components/ui";
import { ToastContainer } from "../components/ToastContainer";
import { NotificationDrawer } from "../components/NotificationDrawer";
import { NotificationService } from "./services/notificationService";
import { apiFetch, openAuth, openErp, openTrialSelectDialog, isTrustedOrigin } from "./api";
import { saveAccount } from "./accountHistory";

const initialUser = () => ({
  name: localStorage.getItem("fa_user_name") || "",
  email: localStorage.getItem("fa_user_email") || "",
  plan: localStorage.getItem("fa_plan") || localStorage.getItem("fa_subscription_plan") || "",
  subscriptionId: localStorage.getItem("fa_subscription_id") || ""
});

export function App() {
  const [view, setView] = useState("loading");
  const [user, setUser] = useState(initialUser);
  const [order, setOrder] = useState(null);
  const orderRef = useRef(null);
  const [toasts, setToasts] = useState([]);
  const [notifications, setNotifications] = useState([]);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [drawerTop, setDrawerTop] = useState("56px");
  const [showTrialPopup, setShowTrialPopup] = useState(false);
  const [busy, setBusy] = useState(false);

  // Fetch notifications history when user email is available
  const loadNotifications = useCallback(async () => {
    const list = await NotificationService.fetchNotifications();
    setNotifications(list);
  }, []);

  useEffect(() => {
    if (user.email) {
      loadNotifications();
    }
  }, [user.email, loadNotifications]);

  const dismissToast = useCallback((id) => {
    setToasts((prev) =>
      prev.map((t) => (t.id === id ? { ...t, hiding: true } : t))
    );
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 200);
  }, []);

  // Universal notify / showStatus function
  // Signature: notify(titleOrMessage, type, detail, provider)
  const notify = useCallback(
    (titleOrMessage, type = "success", detail = "", provider = undefined) => {
      if (!titleOrMessage) return;

      // Detect error from title if not explicitly passed
      let resolvedType = type;
      if (
        type === "success" &&
        (/error|fail|failed|unable|denied|invalid|rejected/i.test(titleOrMessage))
      ) {
        resolvedType = "error";
      }

      const id = Date.now() + "_" + Math.random().toString(36).substring(2, 7);
      const newToast = {
        id,
        title: String(titleOrMessage),
        detail: detail ? String(detail) : "",
        type: resolvedType === "error" ? "error" : "success",
        hiding: false
      };

      // Add to floating toasts (stack up to 4 toasts)
      setToasts((prev) => [newToast, ...prev].slice(0, 4));

      // Auto dismiss toast after 4500ms
      setTimeout(() => {
        dismissToast(id);
      }, 4500);

      // Persist notification to backend & history (skip transient % notifications)
      if (!titleOrMessage.startsWith("Pulling Master Data (") || titleOrMessage.includes("100%")) {
        const tempNotif = {
          id,
          type: resolvedType === "error" ? "error" : "success",
          message: String(titleOrMessage),
          detail: detail ? String(detail) : "",
          provider: provider || null,
          timestamp: new Date().toISOString(),
          read: false
        };
        setNotifications((prev) => [tempNotif, ...prev].slice(0, NotificationService.MAX_ITEMS));

        NotificationService.postNotification(
          resolvedType,
          titleOrMessage,
          detail,
          provider
        ).then((saved) => {
          if (saved) {
            setNotifications((prev) =>
              prev.map((n) => (n.id === id ? saved : n))
            );
          }
        });
      }
    },
    [dismissToast]
  );

  const unreadCount = useMemo(() => {
    return notifications.filter((n) => !n.read).length;
  }, [notifications]);

  const toggleDrawer = useCallback((top = "56px") => {
    setDrawerTop(top);
    setDrawerOpen((prev) => {
      const next = !prev;
      if (next) {
        // Mark all as read
        const unreadIds = notifications.filter((n) => !n.read).map((n) => n.id);
        if (unreadIds.length > 0) {
          NotificationService.markRead(unreadIds);
          setNotifications((curr) => curr.map((n) => ({ ...n, read: true })));
        }
      }
      return next;
    });
  }, [notifications]);

  const clearAllNotifications = useCallback(async () => {
    const ok = await NotificationService.clearAll();
    if (ok) {
      setNotifications([]);
      setDrawerOpen(false);
    } else {
      notify("Couldn't clear notifications. Please try again.", "error");
    }
  }, [notify]);

  // Keep orderRef in sync so asynchronous message callbacks can always read the latest selected plan
  useEffect(() => { orderRef.current = order; }, [order]);

  // Session bootstrap & backend sync
  useEffect(() => {
    let mounted = true;
    async function initSession() {
      const email = localStorage.getItem("fa_user_email");
      const token = localStorage.getItem("fa_jwt_token");
      if (!email) {
        if (mounted) setView("welcome");
        return;
      }
      if (token) {
        try {
          const res = await apiFetch("/api/auth/me");
          const data = await res.json();
          if (mounted && data?.user) {
            const serverPlan = data.user.plan || localStorage.getItem("fa_plan") || localStorage.getItem("fa_subscription_plan") || "";
            const serverSubId = data.user.subscriptionId || localStorage.getItem("fa_subscription_id") || "";
            if (serverPlan) {
              localStorage.setItem("fa_plan", serverPlan);
              localStorage.setItem("fa_subscription_plan", serverPlan);
            }
            if (serverSubId) {
              localStorage.setItem("fa_subscription_id", serverSubId);
            }
            setUser((prev) => ({
              ...prev,
              name: data.user.name || prev.name,
              email: data.user.email || prev.email,
              plan: serverPlan,
              subscriptionId: serverSubId || prev.subscriptionId
            }));
            if (serverPlan) {
              setView("dashboard");
            } else {
              setView("welcome");
              handleOpenTrialSelect();
            }
            return;
          }
        } catch (e) {
          console.warn("Session restore verify failed:", e);
        }
      }
      if (mounted) {
        const storedPlan = localStorage.getItem("fa_plan") || localStorage.getItem("fa_subscription_plan");
        if (storedPlan) {
          setView("dashboard");
        } else {
          setView("welcome");
          handleOpenTrialSelect();
        }
      }
    }
    const timer = setTimeout(initSession, 350);
    return () => { mounted = false; clearTimeout(timer); };
  }, []);

  const handlePaymentSuccess = async (data = {}) => {
    let selectedPlan = data?.plan || orderRef.current?.name || order?.name || "Pro";
    // Normalize casing for backend Joi validation (requires 'Basic', 'Standard', 'Pro')
    if (/^basic$/i.test(selectedPlan)) selectedPlan = "Basic";
    else if (/^standard$/i.test(selectedPlan)) selectedPlan = "Standard";
    else if (/^pro$/i.test(selectedPlan)) selectedPlan = "Pro";

    const email = user.email || localStorage.getItem("fa_user_email");
    if (email) {
      try {
        await apiFetch("/api/payments/complete", {
          method: "POST",
          body: JSON.stringify({ email, plan: selectedPlan })
        });
      } catch (err) {
        console.warn("Could not sync payment complete to backend:", err);
      }
    }

    // Refresh official profile from backend
    let subId = data?.subscriptionId || orderRef.current?.subscriptionId || "";
    let finalPlan = selectedPlan;
    try {
      const meRes = await apiFetch("/api/auth/me");
      const meData = await meRes.json();
      if (meData?.user) {
        finalPlan = meData.user.plan || finalPlan;
        subId = meData.user.subscriptionId || subId;
      }
    } catch (_) {}

    if (!subId) {
      subId = localStorage.getItem("fa_subscription_id") || ("FA-SUB-" + Math.floor(100000 + Math.random() * 900000));
    }

    localStorage.setItem("fa_plan", finalPlan);
    localStorage.setItem("fa_subscription_plan", finalPlan);
    localStorage.setItem("fa_subscription_id", subId);

    setUser((prev) => ({
      ...prev,
      plan: finalPlan,
      subscriptionId: subId
    }));

    notify("Payment successful.", "success", `Subscribed to the ${finalPlan} plan.`);
    setView("success");
  };

  useEffect(() => {
    const receive = (event) => {
      if (!isTrustedOrigin(event.origin)) return;
      let data = event.data;
      if (typeof data === "string") {
        try { data = JSON.parse(data); } catch (_) {}
      }
      if (["payment_success", "checkout_complete"].includes(data?.type)) {
        handlePaymentSuccess(data);
      }
      if (["qb_connected", "xero_connected"].includes(data)) {
        setView("dashboard");
      }
    };
    window.addEventListener("message", receive);
    return () => window.removeEventListener("message", receive);
  }, [user.email]);

  const handleOpenTrialSelect = useCallback(() => {
    setShowTrialPopup(true);
  }, []);

  const saveUser = async (profile) => {
    const next = {
      name: profile.name || profile.user?.name || "",
      email: profile.email || profile.user?.email || "",
      plan: profile.plan || profile.user?.plan || localStorage.getItem("fa_plan") || localStorage.getItem("fa_subscription_plan") || "",
      subscriptionId: profile.subscriptionId || profile.user?.subscriptionId || localStorage.getItem("fa_subscription_id") || ""
    };
    Object.entries({
      fa_user_name: next.name,
      fa_user_email: next.email,
      fa_plan: next.plan,
      fa_subscription_plan: next.plan,
      fa_subscription_id: next.subscriptionId,
      fa_jwt_token: profile.token || "",
      fa_refresh_token: profile.refreshToken || ""
    }).forEach(([key, value]) => value && localStorage.setItem(key, value));

    if (next.email) saveAccount({ name: next.name, email: next.email, provider: profile.provider || "google" });

    // Query backend to confirm whether user already has an active subscription
    try {
      const res = await apiFetch("/api/auth/me");
      const data = await res.json();
      if (data?.user?.plan) {
        next.plan = data.user.plan;
        next.subscriptionId = data.user.subscriptionId || next.subscriptionId;
        localStorage.setItem("fa_plan", next.plan);
        localStorage.setItem("fa_subscription_plan", next.plan);
        if (next.subscriptionId) localStorage.setItem("fa_subscription_id", next.subscriptionId);
      }
    } catch (_) {}

    setUser(next);
    if (next.plan) {
      setView("dashboard");
    } else {
      // Pop up the Choose Your Plan dialog, keep background on welcome or loading
      handleOpenTrialSelect();
    }
  };

  const authenticate = (provider, loginHint) => {
    setBusy(true);
    try {
      openAuth(provider, (profile) => {
        saveUser(profile);
        setBusy(false);
      }, loginHint);
    } catch (error) {
      notify(error.message, "error");
      setBusy(false);
    }
  };

  const startTrial = async () => {
    setBusy(true);
    try {
      const result = await apiFetch("/api/auth/start-trial", { method: "POST" }).then((response) => response.json());
      const subId = result.user?.subscriptionId || ("FA-SUB-" + Math.floor(100000 + Math.random() * 900000));
      saveUser({
        user: result.user,
        plan: "Free Trial",
        subscriptionId: subId
      });
      localStorage.setItem("fa_plan", "Free Trial");
      localStorage.setItem("fa_subscription_plan", "Free Trial");
      localStorage.setItem("fa_subscription_id", subId);
      setUser(prev => ({ ...prev, plan: "Free Trial", subscriptionId: subId }));
      setShowTrialPopup(false);
      notify("Free Trial started successfully!", "success");
      setView("dashboard");
    } catch (error) {
      notify("Couldn't start your free trial. Please try again.", "error");
    } finally {
      setBusy(false);
    }
  };

  const logout = () => {
    [
      "fa_user_name", "fa_user_email", "fa_plan", "fa_subscription_plan",
      "fa_subscription_id", "fa_jwt_token", "fa_refresh_token",
      "fa_erp_connected", "fa_erp_type"
    ].forEach((key) => localStorage.removeItem(key));
    setUser({});
    setNotifications([]);
    setShowTrialPopup(false);
    setView("welcome");
  };

  const content = useMemo(() => {
    if (view === "loading") return <Loading />;
    if (view === "welcome") return <Welcome busy={busy} onAuth={authenticate}/>;
    if (view === "trial") return <TrialSelect busy={busy} onTrial={startTrial} onPlans={() => setView("plans")}/>;
    if (view === "plans") return <Plans user={user} onBack={() => setView(user?.plan ? "dashboard" : "welcome")} onSelect={(plan, price, cycle) => { setOrder({ ...plan, price, cycle }); setView("payment"); }}/>;
    if (view === "payment") return <Payment user={user} order={order} notify={notify} onBack={() => setView("plans")} onDone={handlePaymentSuccess}/>;
    if (view === "success") return <Success user={user} order={order} onContinue={() => setView("dashboard")} />;
    return (
      <Dashboard
        user={user}
        notify={notify}
        unreadCount={unreadCount}
        onToggleNotifications={toggleDrawer}
        onLogout={logout}
        onChangePlan={() => setView("plans")}
        onConnect={(provider) => {
          const popup = openErp(provider, user);
          if (!popup) notify("The connection window was blocked. Please allow popups and try again.", "error");
        }}
      />
    );
  }, [view, user, busy, order, unreadCount, toggleDrawer]);

  return (
    <div className="app">
      <ToastContainer toasts={toasts} onDismiss={dismissToast} />
      {drawerOpen && (
        <NotificationDrawer
          notifications={notifications}
          topOffset={drawerTop}
          onClose={() => setDrawerOpen(false)}
          onClearAll={clearAllNotifications}
        />
      )}
      {showTrialPopup && (
        <TrialSelectModal
          onTrial={startTrial}
          onPlans={() => {
            setShowTrialPopup(false);
            setView("plans");
          }}
          onClose={() => setShowTrialPopup(false)}
          busy={busy}
        />
      )}
      {content}
    </div>
  );
}

