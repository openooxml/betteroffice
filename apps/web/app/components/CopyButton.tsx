"use client";

import { useRef, useState } from "react";

// copies text to the clipboard and flashes a confirmation state
export function CopyButton({ text }: { text: string }) {
  const [copied, setCopied] = useState(false);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);

  async function copy() {
    await navigator.clipboard.writeText(text);
    setCopied(true);
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => setCopied(false), 1600);
  }

  return (
    <button
      type="button"
      className="shrink-0 cursor-pointer appearance-none rounded border border-line bg-transparent px-2.5 py-1.5 font-mono text-[0.625rem] tracking-[0.12em] text-ink uppercase transition-colors hover:border-dim hover:text-fg data-[copied=true]:border-acc/35 data-[copied=true]:text-acc"
      data-copied={copied}
      onClick={copy}
      aria-label={copied ? "Copied" : "Copy to clipboard"}
    >
      {copied ? "copied" : "copy"}
    </button>
  );
}
