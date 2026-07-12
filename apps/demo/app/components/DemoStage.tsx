import Link from "next/link";
import type { Format } from "../../lib/formats";

export function DemoStage({ format }: { format: Format }) {
  const editor = format.status === "live" ? mountEditor(format.id) : null;
  return (
    <div className="px-8 py-16 max-[44rem]:px-5 max-[44rem]:py-12">
      <p className="font-mono text-[0.6875rem] uppercase tracking-[0.16em] text-dim">
        {format.kind}
      </p>
      <h1 className="mt-3 text-[2rem] leading-tight font-semibold tracking-[-0.03em] text-fg">
        {format.name}
      </h1>
      <p className="mt-3 max-w-lg text-ink">{format.tagline}</p>

      {editor ?? (
        <div className="mt-10 flex h-72 items-center justify-center rounded-md border border-line-soft bg-surface">
          <span className="inline-flex items-center gap-2 font-mono text-xs text-dim before:size-1.5 before:rounded-full before:bg-faint">
            coming soon
          </span>
        </div>
      )}

      <Link
        href="/"
        className="mt-8 inline-block font-mono text-xs text-dim no-underline transition-colors hover:text-fg"
      >
        ← all demos
      </Link>
    </div>
  );
}

function mountEditor(_id: Format["id"]): React.ReactNode {
  return null;
}
