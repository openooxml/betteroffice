import type { CSSProperties } from "react";
import { FormatCard } from "../components/FormatCard";
import { listedFormats, previewFormats } from "../../lib/formats";

const label = "font-mono text-[0.6875rem] uppercase tracking-[0.14em] text-dim";

// a compact head in the website's section-head type, so every editor link sits in the first screen
export default function Home() {
  return (
    <>
      <section aria-labelledby="hero">
        <div className="wrap grid items-end gap-x-12 gap-y-5 pt-[clamp(2rem,3.75vw,3.5rem)] pb-[clamp(2rem,3.2vw,3rem)] lg:grid-cols-12">
          <div className="lg:col-span-7">
            <p className={`${label} rise mb-5 sm:mb-7`}>Browser playground</p>
            <h1
              id="hero"
              className="settle -ml-[0.03em] text-[clamp(2.25rem,5.6vw,4.75rem)] leading-[0.94] font-[560] tracking-[-0.045em]"
            >
              Try BetterOffice
            </h1>
          </div>
          <p
            className="rise max-w-[34rem] text-[clamp(1rem,1.3vw,1.125rem)] leading-[1.6] text-ink lg:col-span-5 lg:pb-1"
            style={{ "--i": 1 } as CSSProperties}
          >
            Open real Office files directly in your browser. Parsing, editing,
            calculation, and rendering stay local on our Rust and WebAssembly
            engines.
          </p>
        </div>
      </section>

      <section aria-labelledby="editors">
        <div className="wrap rule pt-[clamp(1.5rem,2.5vw,2rem)] pb-[clamp(3rem,6vw,5rem)]">
          <h2 id="editors" className={`${label} mb-5 flex gap-4`}>
            <span className="tabular-nums">01</span>
            <span className="text-fg">Editors</span>
          </h2>
          <div className="grid gap-px border border-line-soft bg-line-soft sm:grid-cols-2 lg:grid-cols-4">
            {[...listedFormats, ...previewFormats].map((format) => (
              <FormatCard key={format.id} format={format} preview={previewFormats.includes(format)} />
            ))}
          </div>
        </div>
      </section>
    </>
  );
}
