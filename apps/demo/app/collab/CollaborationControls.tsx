"use client";

import { useEffect, useState } from "react";
import { cn } from "../../lib/cn";
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

  const indicator = error ? "error" : status;

  return (
    <div className="inline-flex flex-none items-center gap-2 font-mono text-[11px] leading-none">
      <span
        className="inline-flex min-h-7 items-center gap-1.5 rounded-full border border-hairline-strong bg-white px-[9px] whitespace-nowrap text-mute"
        data-status={indicator}
        role="status"
        title={error ?? undefined}
      >
        <span
          className={cn(
            "size-[7px] rounded-full",
            indicator === "connected"
              ? "bg-acc"
              : indicator === "connecting"
                ? "bg-[#d97706]"
                : "bg-danger",
          )}
          aria-hidden="true"
        />
        {label}
      </span>
      <button
        className="min-h-7 cursor-pointer rounded-[5px] border border-hairline-strong bg-white px-[9px] whitespace-nowrap text-fg hover:bg-surface max-[760px]:hidden"
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
