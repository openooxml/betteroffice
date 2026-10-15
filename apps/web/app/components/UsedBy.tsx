import type { CSSProperties } from "react";
import { USED_BY, USED_BY_LABEL, USED_BY_MORE, USED_BY_ROTATE, type UsedByEntry } from "../used-by";

// the shared x-height every wordmark is sized to: 11px on phones up to 13px from about 40rem
const X = "clamp(11px, calc(8px + 0.8vw), 13px)";
const px = (k: number) => `calc(${X} * ${k})`;

// ticker only: each half of the track holds at least this many names, so it is wider than the band
const MIN_PER_HALF = 8;
const SECONDS_PER_NAME = 5.5;

// an inline-block whose bottom margin edge is the wordmark's baseline, so logos and text
// in one line share it; sized so every wordmark has the same x-height. The name is visually
// hidden text rather than an aria-label, so text extractors read it as well as screen readers.
function Mark({ entry }: { entry: UsedByEntry }) {
  if (!entry.logo) {
    return (
      <span style={{ fontSize: px(1 / 0.53), fontWeight: 600, letterSpacing: "-0.03em" }}>{entry.name}</span>
    );
  }
  const { src, box, baseline, xHeight } = entry.logo;
  const [w, h] = box;
  return (
    <>
      <span
        aria-hidden="true"
        style={{
          display: "inline-block",
          width: px(w / xHeight),
          height: px(h / xHeight),
          marginBottom: px(-(h - baseline) / xHeight),
          backgroundColor: "currentColor",
          mask: `url(${src}) center / contain no-repeat`,
          WebkitMask: `url(${src}) center / contain no-repeat`,
        }}
      />
      <span className="sr-only">{entry.name}</span>
    </>
  );
}

const row: CSSProperties = {
  display: "flex",
  flexWrap: "wrap",
  alignItems: "baseline",
  justifyContent: "space-between",
  columnGap: "clamp(1.5rem, 4vw, 3rem)",
  rowGap: "1.5rem",
  paddingBlock: "1.25rem",
  listStyle: "none",
};

function Row({ entries }: { entries: UsedByEntry[] }) {
  return (
    <ul style={row}>
      {entries.map((entry) => (
        <li key={entry.name}>
          <a href={entry.href} target="_blank" rel="noopener" className="text-dim no-underline transition-colors hover:text-fg">
            <Mark entry={entry} />
          </a>
        </li>
      ))}
      <li className="font-mono text-[0.8125rem] text-dim">
        {USED_BY_MORE}
      </li>
    </ul>
  );
}

// the first copy of the names is the real list, every repeat is aria-hidden and inert,
// and the track slides by exactly one half so the loop is seamless
function Ticker({ entries }: { entries: UsedByEntry[] }) {
  const perHalf = Math.ceil(MIN_PER_HALF / entries.length);
  return (
    <div className="marquee" style={{ "--marquee-s": `${entries.length * perHalf * SECONDS_PER_NAME}s` } as CSSProperties}>
      <div className="marquee-track">
        {Array.from({ length: perHalf * 2 }, (_, copy) => (
          <ul key={copy} className="marquee-copy" aria-hidden={copy > 0 || undefined} inert={copy > 0 || undefined}>
            {entries.map((entry) => (
              <li key={entry.name}>
                <a href={entry.href} target="_blank" rel="noopener">
                  <Mark entry={entry} />
                </a>
              </li>
            ))}
          </ul>
        ))}
      </div>
    </div>
  );
}

export function UsedBy() {
  const entries = USED_BY.filter((entry) => entry.enabled);
  if (entries.length === 0) return null;
  return (
    <section aria-labelledby="used-by">
      <div className="wrap rule grid items-baseline gap-x-12 py-4 md:grid-cols-[12rem_1fr]">
        <h2 id="used-by" className="font-mono text-[0.8125rem] text-dim max-md:pt-5">
          {USED_BY_LABEL}
        </h2>
        {USED_BY_ROTATE ? <Ticker entries={entries} /> : <Row entries={entries} />}
      </div>
    </section>
  );
}
