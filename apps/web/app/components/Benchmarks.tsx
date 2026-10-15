import type { ReactNode } from "react";
import { ArrowDown, ArrowUp } from "lucide-react";
import { fidelityCharts, sharedVersion, type FidelityBar, type FidelityChart, type FidelityFormat } from "../../lib/fidelity";
import type { AgentEvalCard } from "../../lib/agent-evals";

const mono = "font-mono text-[0.6875rem]";

function Version({ commit, commitUrl, libreoffice }: { commit: string; commitUrl: string; libreoffice: string }) {
  return (
    <p className={`${mono} text-dim`}>
      BetterOffice{" "}
      <a href={commitUrl} className="text-ink underline decoration-line underline-offset-2 hover:text-fg">
        {commit}
      </a>
      , LibreOffice {libreoffice}
    </p>
  );
}

// both engines on one scale; every bar's engine and value are text, the bar itself is decoration
function Chart({ chart }: { chart: FidelityChart }) {
  const Cue = chart.better === "higher" ? ArrowUp : ArrowDown;
  return (
    <div className="grid gap-3.5">
      <div>
        <div className="flex items-baseline justify-between gap-3">
          <h4 className="text-sm leading-snug font-[560] tracking-[-0.01em]">{chart.label}</h4>
          <span className={`${mono} inline-flex shrink-0 items-center gap-1 text-dim`}>
            <Cue size={10} strokeWidth={2} aria-hidden="true" />
            {chart.better} is better
          </span>
        </div>
        <span className="text-xs text-dim">{chart.hint}</span>
      </div>
      <Bars bars={chart.bars} labelColumn="grid-cols-[5.5rem_1fr]" />
    </div>
  );
}

function Bars({ bars, labelColumn }: { bars: FidelityBar[]; labelColumn: string }) {
  return (
    <dl className="grid gap-2">
      {bars.map((bar) => (
        <div key={bar.who} className={`grid ${labelColumn} items-center gap-2.5`}>
          <dt className={`${mono} ${bar.ours ? "text-fg" : "text-dim"}`}>{bar.who}</dt>
          <dd className="grid grid-cols-[1fr_3.25rem] items-center gap-2.5">
            <span aria-hidden="true" className="h-2 bg-line-soft">
              <span
                className={`block h-full ${bar.ours ? "bg-acc" : "bg-dim/55"}`}
                style={{ width: `${(bar.share * 100).toFixed(2)}%` }}
              />
            </span>
            <span className={`${mono} text-right tabular-nums ${bar.ours ? "text-fg" : "text-ink"}`}>{bar.text}</span>
          </dd>
        </div>
      ))}
    </dl>
  );
}

export function AgentEditCards({ cards, names }: { cards: AgentEvalCard[]; names: Record<string, string> }) {
  return (
    <ul className="grid list-none gap-px border border-line-soft bg-line-soft lg:grid-cols-3">
      {cards.map((card) => (
        <li key={card.format} className="grid content-start gap-5 bg-bg p-7 max-sm:p-6">
          <div className="flex items-baseline justify-between gap-4">
            <h4 className="text-sm leading-snug font-[560] tracking-[-0.01em]">
              {names[card.format] ?? card.format.toUpperCase()}
            </h4>
            <span className={`${mono} text-dim`}>.{card.format}</span>
          </div>
          <Bars bars={card.bars} labelColumn="grid-cols-[8.25rem_1fr]" />
        </li>
      ))}
    </ul>
  );
}

function Charts({ charts, className }: { charts: FidelityChart[]; className: string }) {
  if (charts.length === 0) return null;
  return (
    <div className={`grid content-start gap-9 ${className}`}>
      {charts.map((chart) => (
        <Chart key={chart.row} chart={chart} />
      ))}
    </div>
  );
}

const inGroup = (f: FidelityFormat, group: FidelityChart["group"]) => fidelityCharts(f).filter((c) => c.group === group);

// one card per format: quality bars, then time bars on their own row (a shared subgrid row
// across the cards at lg), every published row verbatim behind "All results"
export function BenchmarkCards({
  formats,
  names,
  children,
}: {
  formats: FidelityFormat[];
  names: Record<string, string>;
  children?: ReactNode;
}) {
  const shared = sharedVersion(formats);
  return (
    <>
      <ul className="grid list-none gap-px border border-line-soft bg-line-soft lg:grid-cols-3">
        {formats.map((f) => (
          <li
            key={f.format}
            className="flex flex-col gap-9 bg-bg p-7 max-sm:p-6 lg:row-span-4 lg:grid lg:grid-rows-subgrid"
          >
            <div className="flex items-baseline justify-between gap-4">
              <h3 className="text-[1.0625rem] leading-snug font-[560] tracking-[-0.015em]">
                {names[f.format] ?? f.format.toUpperCase()}
              </h3>
              <span className={`${mono} text-dim`}>.{f.format}</span>
            </div>
            <Charts charts={inGroup(f, "quality")} className="lg:row-start-2" />
            <Charts charts={inGroup(f, "time")} className="border-t border-line-soft pt-9 lg:row-start-3" />
            <div className="mt-auto grid gap-4 lg:row-start-4">
              <details className="group border-t border-line-soft pt-4">
                <summary
                  className={`${mono} flex w-fit cursor-pointer list-none items-center gap-2 text-dim transition-colors hover:text-fg [&::-webkit-details-marker]:hidden`}
                >
                  <span aria-hidden="true" className="transition-transform group-open:rotate-45">
                    +
                  </span>
                  All results
                </summary>
                <table className={`${mono} mt-3 w-full border-collapse tabular-nums`}>
                  <caption className="sr-only">
                    {f.format.toUpperCase()} results, BetterOffice {f.commit} against LibreOffice {f.libreoffice}
                  </caption>
                  <thead>
                    <tr className="border-b border-line text-dim">
                      <th scope="col" className="py-2 text-left font-normal">
                        <span className="sr-only">Metric</span>
                      </th>
                      <th scope="col" className="py-2 pl-3 text-right font-normal text-fg">
                        BetterOffice
                      </th>
                      <th scope="col" className="py-2 pl-3 text-right font-normal">
                        LibreOffice
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {f.rows.map((row) => (
                      <tr key={row.label} className="border-b border-line-soft">
                        <th scope="row" className="py-2 text-left font-normal text-dim">
                          {row.label}
                        </th>
                        <td className="py-2 pl-3 text-right text-fg">{row.ours}</td>
                        <td className="py-2 pl-3 text-right text-dim">{row.theirs}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </details>
              {!shared && <Version {...f} />}
            </div>
          </li>
        ))}
      </ul>
      <div className="mt-6 flex flex-wrap items-baseline justify-between gap-x-8 gap-y-3">
        {shared && <Version {...shared} />}
        {children}
      </div>
    </>
  );
}
