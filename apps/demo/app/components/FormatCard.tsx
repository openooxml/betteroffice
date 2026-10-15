import Link from "next/link";
import { FormatGlyph } from "../../../../shared/format-glyph";
import { PREVIEW_STATUS, type Format } from "../../lib/formats";

// one editor on the launcher: its line drawing, then name, status, kind and tagline;
// a preview format is listed without a link
export function FormatCard({ format, preview = false }: { format: Format; preview?: boolean }) {
  const live = format.status === "live" && !preview;
  const content = (
    <>
      <div className="relative aspect-[2/1] border-b border-line-soft bg-surface text-fg transition-colors group-hover:bg-[color-mix(in_oklab,var(--color-surface),var(--color-fg)_4%)] max-sm:aspect-auto max-sm:border-r max-sm:border-b-0">
        <div className="absolute inset-x-3 top-3 bottom-1 max-sm:inset-x-2.5 max-sm:inset-y-4">
          <FormatGlyph format={format.id} />
        </div>
      </div>
      <div className="flex flex-1 flex-col gap-2 p-5 xl:p-6">
        <div className="flex items-center justify-between gap-3">
          <h3 className="text-[1.0625rem] leading-snug font-[560] tracking-[-0.015em]">{format.name}</h3>
          <Status status={preview ? PREVIEW_STATUS : live ? "available" : "coming"} />
        </div>
        <p className="font-mono text-[0.6875rem] uppercase tracking-[0.14em] text-dim">{format.kind}</p>
        <p className="text-sm leading-[1.65] text-ink">{format.tagline}</p>
        {!preview && (
          <span className="mt-auto pt-3 font-mono text-xs text-dim transition-colors group-hover:text-fg">
            {live ? "open demo →" : "coming soon"}
          </span>
        )}
      </div>
    </>
  );
  const card = "flex flex-col bg-bg max-sm:grid max-sm:grid-cols-[8.5rem_1fr]";

  if (!live) return <div className={card}>{content}</div>;

  return (
    <Link href={`/${format.id}`} className={`group ${card} no-underline`}>
      {content}
    </Link>
  );
}

function Status({ status }: { status: string }) {
  const available = status === "available";
  return (
    <span
      className={`inline-flex shrink-0 items-center gap-1.5 border px-1.5 py-px font-mono text-[0.625rem] uppercase tracking-[0.12em] whitespace-nowrap ${
        available ? "border-acc/40 text-acc" : "border-dashed border-line text-dim"
      }`}
    >
      {available && <span className="size-1.5 bg-acc" />}
      {status}
    </span>
  );
}
