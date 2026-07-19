"use client";

import { useEffect, useState } from "react";
import type { CollaborationStatus } from "./types";

interface CollaborationControlsProps {
  status: CollaborationStatus;
  synced: boolean;
  peerCount: number | null;
  error: string | null;
}

export function CollaborationControls({
  status,
  synced,
  peerCount,
  error,
}: CollaborationControlsProps) {
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    if (!copied) return;
    const timer = setTimeout(() => setCopied(false), 1800);
    return () => clearTimeout(timer);
  }, [copied]);

  const label = [
    error ? "error" : status,
    synced ? "synced" : null,
    peerCount === null
      ? null
      : `${peerCount} ${peerCount === 1 ? "peer" : "peers"}`,
  ]
    .filter(Boolean)
    .join(" · ");

  return (
    <div className="collab-controls">
      <span
        className="collab-status"
        data-status={error ? "error" : status}
        role="status"
        title={error ?? undefined}
      >
        <span className="collab-status-dot" aria-hidden="true" />
        {label}
      </span>
      <button
        className="collab-copy"
        type="button"
        onClick={() => {
          void navigator.clipboard
            .writeText(window.location.href)
            .then(() => setCopied(true));
        }}
      >
        {copied ? "Copied" : "Copy collaboration link"}
      </button>
    </div>
  );
}
