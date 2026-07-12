import Link from "next/link";
import type { Format } from "../../lib/formats";

export function DemoStage({ format }: { format: Format }) {
  const editor = format.status === "live" ? mountEditor(format.id) : null;
  return (
    <div>
      <div className="font-mono text-[0.7rem] uppercase tracking-[0.16em] text-faint">
        {format.kind}
      </div>
      <h1 className="mt-2 text-2xl font-medium tracking-tight text-fg">
        {format.name} demo
      </h1>
      <p className="mt-3 max-w-lg text-sm leading-relaxed text-ink">
        {format.tagline}
      </p>

      {editor ?? (
        <div className="mt-10 flex h-72 items-center justify-center rounded-lg border border-dashed border-line bg-surface">
          <span className="font-mono text-xs text-faint">Coming soon</span>
        </div>
      )}

      <Link
        href="/"
        className="mt-8 inline-block font-mono text-xs text-ink no-underline transition-colors hover:text-fg"
      >
        ← All demos
      </Link>
    </div>
  );
}

// Each format's client-only editor is dynamic()-mounted here once its package lands.
function mountEditor(_id: Format["id"]): React.ReactNode {
  return null;
}
