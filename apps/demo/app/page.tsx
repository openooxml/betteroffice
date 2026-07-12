import { FormatCard } from "./components/FormatCard";
import { formats } from "../lib/formats";

export default function Home() {
  return (
    <>
      <section className="relative px-8 pt-24 pb-16 max-[44rem]:px-5 max-[44rem]:pt-16 max-[44rem]:pb-12">
        <p className="mb-6 font-mono text-[0.6875rem] uppercase tracking-[0.16em] text-dim">
          Browser playground
        </p>
        <h1 className="mb-4 text-[clamp(2.5rem,8vw,3.75rem)] leading-[1.05] font-[650] tracking-[-0.035em] text-fg">
          Try BetterOffice
        </h1>
        <p className="max-w-[32rem] text-lg text-ink">
          Open real Office files directly in your browser. Parsing, editing,
          calculation, and rendering stay local on our Rust and WebAssembly
          engines.
        </p>
      </section>

      <section className="sec relative border-t border-line-soft px-8 py-16 max-[44rem]:px-5 max-[44rem]:py-12">
        <p className="mb-8 flex items-baseline gap-3 font-mono text-[0.6875rem] uppercase tracking-[0.16em] text-dim">
          <span className="text-faint">01</span> Editors
        </p>
        <div className="grid grid-cols-2 gap-px overflow-hidden rounded-md border border-line-soft bg-line-soft max-[44rem]:grid-cols-1">
          {formats.map((f) => (
            <FormatCard key={f.id} format={f} />
          ))}
        </div>
      </section>
    </>
  );
}
