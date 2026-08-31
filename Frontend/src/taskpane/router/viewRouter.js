/**
 * Lightweight internal view controller.
 *
 * Views are plain `.view` sections in taskpane.html; exactly one carries the
 * `active` class at a time. Modals are separate overlay elements toggled
 * through `display`, kept here so view/modal visibility lives in one place.
 */
import { AppState } from "../state/appState.js";

const ViewRouter = {
    /**
     * Shows a view by its ID name (e.g. "Welcome", "Dashboard")
     * @param {string} name - View name that maps to #view<Name>
     */
    show(name) {
        document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
        const el = document.getElementById("view" + name);
        if (el) el.classList.add("active");
        if (name && name !== "Loading" && name !== "Error") {
            localStorage.setItem("fa_last_view", name);
        }
        if (name === "Plans") {
            const nameVal = AppState.userName || AppState.userEmail || "User";
            const initial = nameVal.charAt(0).toUpperCase();
            const avatarEl = document.getElementById("plansUserAvatar");
            const emailEl = document.getElementById("plansUserEmail");
            if (avatarEl) avatarEl.textContent = initial;
            if (emailEl) emailEl.textContent = AppState.userEmail || "";
        }
    },

    /** Returns the currently active view name (e.g. "Dashboard", "Plans", "Payment"). */
    getCurrentView() {
        const activeView = document.querySelector(".view.active");
        if (activeView && activeView.id) {
            return activeView.id.replace("view", "");
        }
        return localStorage.getItem("fa_last_view") || "";
    },

    /**
     * Opens a modal overlay by element id. No-ops when the element is absent,
     * matching the optional-chaining style used by every call site.
     */
    openModal(id) {
        const modal = document.getElementById(id);
        if (modal) modal.style.display = "flex";
    },

    /** Closes a modal overlay by element id. */
    closeModal(id) {
        const modal = document.getElementById(id);
        if (modal) modal.style.display = "none";
    },

    /** True when the modal overlay is currently on screen. */
    isModalOpen(id) {
        const modal = document.getElementById(id);
        return !!modal && modal.style.display === "flex";
    }
};

export { ViewRouter };
