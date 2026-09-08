import React, { useEffect } from "react";

export function CompanyContextMenu({
  visible,
  x,
  y,
  company,
  onEdit,
  onDisconnect,
  onClose,
}) {
  useEffect(() => {
    const handleClick = () => {
      if (visible) onClose();
    };
    window.addEventListener("click", handleClick);
    return () => window.removeEventListener("click", handleClick);
  }, [visible, onClose]);

  if (!visible) return null;

  return (
    <div
      className="fa-context-menu"
      style={{
        display: "block",
        top: y,
        left: x,
      }}
      onClick={(e) => e.stopPropagation()}
    >
      <button
        className="fa-context-item"
        onClick={() => {
          onClose();
          if (company) onEdit(company);
        }}
      >
        Edit
      </button>
      <button
        className="fa-context-item"
        onClick={() => {
          onClose();
          if (company) onDisconnect(company);
        }}
      >
        Disconnect
      </button>
    </div>
  );
}
