"use client";

import { useEffect, useState } from "react";
import { cn } from "../../lib/cn";
import { barBordered } from "../components/DemoTopBar";
import type { CollaborationStatus } from "./types";

interface CollaborationControlsProps {
  status: CollaborationStatus;
  synced: boolean;
  peerCount: number | null;
  error: string | null;
  /** False for a locally opened document, which no link can reach. */
  shared?: boolean;
}

export function CollaborationControls({
  status,
  synced,
  peerCount,
  error,
  shared = true,
}: CollaborationControlsProps) {
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    if (!copied) return;
    const timer = setTimeout(() => setCopied(false), 1800);
    return () => clearTimeout(timer);
  }, [copied]);

  const label = shared
    ? [
        error ? "error" : status,
        synced ? "synced" : null,
        peerCount === null
          ? null
          : `${peerCount} ${peerCount === 1 ? "peer" : "peers"}`,
      ]
        .filter(Boolean)
        .join(" · ")
    : "not shared";

  const indicator = shared ? (error ? "error" : status) : "disconnected";

  return (
    <div className="inline-flex flex-none items-center gap-1.5 text-[12.5px] leading-none">
      <span
        className="inline-flex h-7 items-center gap-1.5 rounded-full border border-[rgba(31,31,31,0.16)] bg-white px-2.5 whitespace-nowrap text-[#444746]"
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
      {shared && (
        <button
          className={`${barBordered} max-[760px]:hidden`}
          type="button"
          onClick={() => {
            void navigator.clipboard
              .writeText(window.location.href)
              .then(() => setCopied(true));
          }}
        >
          {copied ? "Copied" : "Copy collaboration link"}
        </button>
      )}
    </div>
  );
}
