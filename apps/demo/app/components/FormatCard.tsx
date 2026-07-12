import Link from "next/link";
import type { Format } from "../../lib/formats";

export function FormatCard({ format }: { format: Format }) {
  const live = format.status === "live";
  return (
    <Link
      href={`/${format.id}`}
      className="group flex flex-col justify-between rounded-lg border border-line bg-surface p-6 no-underline transition-colors hover:border-fg"
    >
      <div>
        <div className="flex items-center justify-between">
          <span className="font-mono text-lg text-fg">{format.name}</span>
          <StatusBadge live={live} />
        </div>
        <p className="mt-1 font-mono text-[0.7rem] uppercase tracking-[0.16em] text-faint">
          {format.kind}
        </p>
        <p className="mt-4 text-sm leading-relaxed text-ink">{format.tagline}</p>
      </div>
      <span className="mt-8 font-mono text-xs text-ink transition-colors group-hover:text-fg">
        {live ? "Open demo →" : "Preview →"}
      </span>
    </Link>
  );
}

function StatusBadge({ live }: { live: boolean }) {
  return (
    <span
      className={
        "rounded-full border px-2 py-0.5 font-mono text-[0.625rem] uppercase tracking-[0.12em] " +
        (live
          ? "border-acc/40 text-acc"
          : "border-line text-faint")
      }
    >
      {live ? "Live" : "Soon"}
    </span>
  );
}
