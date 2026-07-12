import Link from "next/link";
import type { Format } from "../../lib/formats";

export function FormatCard({ format }: { format: Format }) {
  const live = format.status === "live";
  return (
    <Link
      href={`/${format.id}`}
      className="group flex min-h-52 flex-col bg-bg p-6 no-underline transition-colors hover:bg-surface"
    >
      <div>
        <div className="flex items-center justify-between">
          <span className="font-mono text-[0.8125rem] text-fg">
            {format.name}
          </span>
          <StatusBadge live={live} />
        </div>
        <p className="mt-2 font-mono text-[0.625rem] uppercase tracking-[0.14em] text-dim">
          {format.kind}
        </p>
        <p className="mt-4 text-[0.8125rem] leading-relaxed text-ink">
          {format.tagline}
        </p>
      </div>
      <span className="mt-auto pt-6 font-mono text-xs text-dim transition-colors group-hover:text-fg">
        {live ? "open demo →" : "preview →"}
      </span>
    </Link>
  );
}

function StatusBadge({ live }: { live: boolean }) {
  return (
    <span
      className={
        "inline-flex items-center gap-1.5 font-mono text-[0.625rem] uppercase tracking-[0.12em] " +
        (live
          ? "text-acc before:size-1.5 before:rounded-full before:bg-acc before:shadow-[0_0_8px_rgba(5,150,105,0.45)]"
          : "text-dim before:size-1.5 before:rounded-full before:bg-faint")
      }
    >
      {live ? "available" : "coming"}
    </span>
  );
}
