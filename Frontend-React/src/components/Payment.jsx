import { useState } from "react";
import { API_BASE } from "../taskpane/api";
import { Header } from "./ui";

export function Payment({ user, order, onBack, onDone, notify }) {
  const [busy, setBusy] = useState(false);
  const currentOrder = order || { name: "Pro", price: 1999, cycle: "monthly" };
  const proceed = () => {
    setBusy(true);
    const params = new URLSearchParams({
      plan: currentOrder.name,
      price: String(currentOrder.price),
      cycle: currentOrder.cycle,
      email: user.email || localStorage.getItem("fa_user_email") || "",
      token: localStorage.getItem("fa_jwt_token") || "",
    });
    const popup = window.open(
      `${API_BASE}/api/payments/checkout?${params}`,
      "fa_checkout",
      "width=700,height=700",
    );
    if (!popup)
      notify("Checkout popup was blocked. Please allow popups and try again.");
    setBusy(false);
  };
  return (
    <section className="view active">
      <Header title="Secure Checkout" onBack={onBack} user={user} />
      <div className="view-content scrollable">
        <div className="order-summary-card">
          <div className="order-summary-row">
            <span className="order-summary-label">Selected plan</span>
            <span className="order-summary-value">
              {currentOrder.name} Plan
            </span>
          </div>
          <div className="order-summary-row">
            <span className="order-summary-label">Billing cycle</span>
            <span
              className="order-summary-value"
              style={{ textTransform: "capitalize" }}
            >
              {currentOrder.cycle}
            </span>
          </div>
          <div className="order-summary-divider" />
          <div className="order-summary-row total-row">
            <span className="order-summary-label">Total</span>
            <span className="order-summary-total">
              ₹{currentOrder.price.toLocaleString("en-IN")}
            </span>
          </div>
        </div>
        <div className="payment-redirect-box">
          <div className="payment-redirect-icon">🔒</div>
          <h3 className="payment-redirect-title">Secure Payment Processing</h3>
          <p className="payment-redirect-desc">
            You’ll be redirected to our secure payment provider to complete your
            purchase.
          </p>
          <button
            className="payment-proceed-btn"
            disabled={busy}
            onClick={proceed}
          >
            {busy ? "Opening Secure Checkout…" : "Open Secure Checkout"}
          </button>
          <p className="payment-already-text">
            Already paid?{" "}
            <button
              className="link-button"
              onClick={onDone}
              style={{
                background: "none",
                border: "none",
                color: "#2459dd",
                fontWeight: 600,
                cursor: "pointer",
                textDecoration: "underline",
                padding: 0,
                font: "inherit",
              }}
            >
              Verify my payment →
            </button>
          </p>
        </div>
      </div>
    </section>
  );
}
