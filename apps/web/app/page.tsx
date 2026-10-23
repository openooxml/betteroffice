import type { Metadata } from "next";
import type { CSSProperties, ReactNode } from "react";
import { ArrowRight, ArrowUpRight } from "lucide-react";
import { publishedFidelity } from "../lib/fidelity";
import { agentEvalCards, agentEvalRuns, publishedAgentEvals } from "../lib/agent-evals";
import { AgentEditCards, BenchmarkCards } from "./components/Benchmarks";
import { CoEdit } from "./components/CoEdit";
import { CopyButton } from "./components/CopyButton";
import { FormatGlyph } from "./components/FormatGlyph";
import { HeroBackdrop } from "./components/HeroBackdrop";
import { Logo } from "./components/Logo";
import { UsedBy } from "./components/UsedBy";
import {
  AGENT_EVALS,
  CAPABILITIES,
  BENCHMARKS_SECTION,
  CODA,
  BENCHMARKS_SITE,
  COLLABORATION,
  CRATES,
  DEMO,
  DOCS,
  ECOSYSTEMS,
  EDITORS,
  editorOrder,
  FEATURES_LINK,
  FOUNDATION,
  HERO,
  NPM,
  OPENOOXML,
  PACKAGES,
  PACKAGES_SECTION,
  PEERS,
  PYPI,
  REPO,
  RELEASES,
  SITE,
  SITE_DESCRIPTION,
  SUITE,
} from "./content";
import { FEED_URL, publishedPosts } from "../lib/blog";

export const metadata: Metadata = {
  alternates: {
    canonical: "/",
    types: {
      "text/markdown": "/index.md",
      ...(publishedPosts().length > 0 ? { "application/atom+xml": FEED_URL } : {}),
    },
  },
};

const label = "font-mono text-[0.6875rem] uppercase tracking-[0.14em] text-dim";
const h2 =
  "text-[clamp(2.25rem,5.6vw,4.75rem)] leading-[0.94] font-[560] tracking-[-0.045em]";
const lead = "text-[clamp(1rem,1.3vw,1.125rem)] leading-[1.6] text-ink";
const small = "text-sm leading-[1.65] text-ink";
const itemName = "text-[1.0625rem] leading-snug font-[560] tracking-[-0.015em]";
const bodyLink =
  "text-fg underline decoration-faint underline-offset-[3px] transition-colors hover:decoration-fg";
const btn =
  "inline-flex h-11 items-center gap-2.5 border px-5 font-mono text-xs no-underline transition-colors";
const btnPrimary = `${btn} border-fg bg-fg text-bg hover:bg-transparent hover:text-fg`;
const btnGhost = `${btn} border-line text-fg hover:border-fg`;
const miniLink =
  "inline-flex items-center gap-1 font-mono text-[0.6875rem] text-dim no-underline transition-colors hover:text-fg";
const cells = "grid gap-px border border-line-soft bg-line-soft";
const pad = "py-[clamp(4.5rem,10vw,8.5rem)]";

const FIDELITY = publishedFidelity().sort((a, b) => editorOrder(a.format) - editorOrder(b.format));
const EVALS = publishedAgentEvals();
const EVAL_CARDS = agentEvalCards(EVALS).sort((a, b) => editorOrder(a.format) - editorOrder(b.format));
const EDITOR_NAMES = Object.fromEntries(EDITORS.map((e) => [e.format, e.name]));
const BENCH = FIDELITY.length > 0 ? 1 : 0;
const SECTIONS = 4 + BENCH;

const JSON_LD = {
  "@context": "https://schema.org",
  "@type": "SoftwareApplication",
  name: "BetterOffice",
  url: SITE,
  image: `${SITE}/logo.svg`,
  applicationCategory: "BusinessApplication",
  operatingSystem: "Web",
  description: SITE_DESCRIPTION,
  offers: {
    "@type": "Offer",
    price: "0",
    priceCurrency: "USD",
  },
  license: "https://www.apache.org/licenses/LICENSE-2.0",
  softwareHelp: DOCS,
  creator: {
    "@type": "Organization",
    name: "OpenOOXML",
    url: OPENOOXML,
  },
  sameAs: [REPO, NPM, CRATES, PYPI, OPENOOXML],
};

function SectionHead({
  n,
  id,
  name,
  heading,
  children,
}: {
  n: number;
  id: string;
  name: string;
  heading: ReactNode;
  children: ReactNode;
}) {
  const echoes = typeof heading === "string" && heading.toLowerCase().includes(name.toLowerCase());
  return (
    <div className="grid items-end gap-x-12 gap-y-7 lg:grid-cols-12">
      <div className="lg:col-span-7">
        <p className={`${label} mb-7 flex gap-4`}>
          <span className="tabular-nums">
            {String(n).padStart(2, "0")} / {String(SECTIONS).padStart(2, "0")}
          </span>
          {!echoes && <span className="text-fg">{name}</span>}
        </p>
        <h2 id={id} className={h2}>
          {heading}
        </h2>
      </div>
      <div className="flex flex-col gap-4 lg:col-span-5 lg:pb-1.5">{children}</div>
    </div>
  );
}

function Status({ status }: { status: string }) {
  const live = status === "available";
  return (
    <span
      className={`inline-flex shrink-0 items-center gap-1.5 border px-1.5 py-px font-mono text-[0.625rem] uppercase tracking-[0.12em] whitespace-nowrap ${
        live ? "border-acc/40 text-acc" : "border-dashed border-line text-dim"
      }`}
    >
      {live && <span className="size-1.5 bg-acc" />}
      {status}
    </span>
  );
}

export default function Home() {
  return (
    <>
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(JSON_LD) }}
      />

      <section className="relative isolate overflow-hidden" aria-labelledby="hero">
        <HeroBackdrop />
        <div className="wrap flex min-h-[min(calc(100svh-4rem),58rem)] flex-col justify-end pt-28 pb-[clamp(2.5rem,7vh,4.5rem)] max-md:min-h-[min(calc(100svh-4rem),46rem)]">
          <div data-hero-clear="">
            <h1
              id="hero"
              className="settle -ml-[0.045em] text-[clamp(3.25rem,14.6vw,13rem)] leading-[0.8] max-sm:text-[16.2vw] font-[600] tracking-[-0.064em]"
            >
              {HERO.title}
            </h1>
            <div className="mt-[clamp(2rem,4vw,3.5rem)] grid items-end gap-x-12 gap-y-10 lg:grid-cols-12">
              <div className="lg:col-span-7">
                <p className={`${lead} rise max-w-[34rem] text-balance`} style={{ "--i": 1 } as CSSProperties}>
                  {HERO.tagline}
                </p>
                <div
                  className="rise mt-9 grid grid-cols-2 gap-3 sm:flex sm:flex-wrap [&>a:first-child]:col-span-2 [&>a]:justify-center"
                  style={{ "--i": 2 } as CSSProperties}
                >
                  <a href={DEMO} className={btnPrimary}>
                    Try the demo <ArrowRight size={14} strokeWidth={2} />
                  </a>
                  <a href={DOCS} className={btnGhost}>
                    Documentation <ArrowUpRight size={14} strokeWidth={2} />
                  </a>
                  <a href={REPO} target="_blank" rel="noopener" className={btnGhost}>
                    GitHub <ArrowUpRight size={14} strokeWidth={2} />
                  </a>
                </div>
              </div>
              <ul className="list-none border-t border-line lg:col-span-4 lg:col-start-9 max-md:hidden">
                {EDITORS.map((editor) => (
                  <li
                    key={editor.format}
                    className="grid grid-cols-[3.75rem_1fr_auto] items-center gap-3 border-b border-line py-2.5 font-mono text-xs"
                  >
                    <span className="text-dim">.{editor.format}</span>
                    {editor.demo ? (
                      <a href={`${DEMO}/${editor.format}`} className="text-fg no-underline hover:underline">
                        {editor.name}
                      </a>
                    ) : (
                      <span className="text-fg">{editor.name}</span>
                    )}
                    <Status status={editor.status} />
                  </li>
                ))}
              </ul>
            </div>
          </div>
        </div>
      </section>

      <UsedBy />

      <section aria-labelledby="suite">
        <div className={`wrap rule ${pad}`}>
          <SectionHead n={1} id="suite" name={SUITE.label} heading={SUITE.heading}>
            <p className={lead}>{SUITE.prose}</p>
          </SectionHead>
          <div className={`${cells} mt-[clamp(3rem,6vw,4.5rem)] sm:grid-cols-2 lg:grid-cols-4`}>
            {EDITORS.map((editor) => (
              <article className="flex flex-col bg-bg" key={editor.format}>
                <div className="aspect-[4/3] min-h-0 border-b border-line-soft bg-surface px-3 pt-3 text-fg">
                  <FormatGlyph format={editor.format} />
                </div>
                <div className="flex flex-1 flex-col gap-3 p-6">
                  <div className="flex items-center justify-between gap-3">
                    <h3 className={itemName}>{editor.name}</h3>
                    <Status status={editor.status} />
                  </div>
                  <p className={small}>{editor.desc}</p>
                  <div className="mt-auto flex items-center justify-between gap-4 pt-4">
                    <span className="font-mono text-[0.6875rem] text-dim">.{editor.format}</span>
                    {editor.demo && (
                      <a href={`${DEMO}/${editor.format}`} className={miniLink}>
                        demo <ArrowUpRight size={11} strokeWidth={2} />
                      </a>
                    )}
                  </div>
                </div>
              </article>
            ))}
          </div>
        </div>
      </section>

      <section aria-labelledby="foundation">
        <div className={`wrap rule ${pad}`}>
          <SectionHead n={2} id="foundation" name={FOUNDATION.label} heading={FOUNDATION.heading}>
            <p className={lead}>{FOUNDATION.prose}</p>
          </SectionHead>
          <ol className={`${cells} mt-[clamp(3rem,6vw,4.5rem)] list-none sm:grid-cols-2 lg:grid-cols-3`}>
            {CAPABILITIES.map((cap, i) => (
              <li className="flex flex-col gap-3 bg-bg p-7 max-sm:p-6" key={cap.name}>
                <span className="font-mono text-[0.6875rem] text-dim tabular-nums">
                  {String(i + 1).padStart(2, "0")}
                </span>
                <h3 className={`${itemName} mt-6 max-sm:mt-3`}>{cap.name}</h3>
                <p className={small}>{cap.desc}</p>
              </li>
            ))}
          </ol>
          <p className="mt-8">
            <a href={FEATURES_LINK.href} className={miniLink}>
              {FEATURES_LINK.label} <ArrowUpRight size={11} strokeWidth={2} />
            </a>
          </p>
        </div>
      </section>

      {BENCH > 0 && (
        <section aria-labelledby="benchmarks">
          <div className={`wrap rule ${pad}`}>
            <SectionHead
              n={3}
              id="benchmarks"
              name={BENCHMARKS_SECTION.label}
              heading={BENCHMARKS_SECTION.heading}
            >
              <p className={small}>
                Generated by the{" "}
                <a href={BENCHMARKS_SECTION.workflow} className={bodyLink}>
                  Benchmarks action
                </a>
                ; per-document results are on the{" "}
                <a href={BENCHMARKS_SITE} className={bodyLink}>benchmarks site</a>. For how scores,
                timings and coverage are measured, see the{" "}
                <a href={BENCHMARKS_SECTION.methodology} className={bodyLink}>
                  benchmark methodology
                </a>
                .
              </p>
            </SectionHead>
            <div className="mt-[clamp(3rem,6vw,4.5rem)]">
              <BenchmarkCards formats={FIDELITY} names={EDITOR_NAMES}>
                <a href={BENCHMARKS_SITE} className={miniLink}>
                  {BENCHMARKS_SECTION.more} <ArrowUpRight size={11} strokeWidth={2} />
                </a>
              </BenchmarkCards>
            </div>
            {EVAL_CARDS.length > 0 && (
              <>
                <div className="mt-[clamp(3.5rem,7vw,5.5rem)] grid items-end gap-x-12 gap-y-4 lg:grid-cols-12">
                  <h3 className="text-[clamp(1.375rem,2.4vw,1.75rem)] leading-[1.1] font-[560] tracking-[-0.03em] lg:col-span-7">
                    {AGENT_EVALS.heading}
                  </h3>
                  <p className={`${small} lg:col-span-5`}>{AGENT_EVALS.lead(agentEvalRuns(EVALS))}</p>
                </div>
                <div className="mt-8">
                  <AgentEditCards cards={EVAL_CARDS} names={EDITOR_NAMES} />
                </div>
                <p className="mt-6">
                  <a href={AGENT_EVALS.method} className={miniLink}>
                    {AGENT_EVALS.more} <ArrowUpRight size={11} strokeWidth={2} />
                  </a>
                </p>
              </>
            )}
          </div>
        </section>
      )}

      <section aria-labelledby="collaboration">
        <div className={`wrap rule ${pad}`}>
          <SectionHead
            n={3 + BENCH}
            id="collaboration"
            name={COLLABORATION.label}
            heading={COLLABORATION.heading}
          >
            <p className={lead}>{COLLABORATION.prose}</p>
            <p className={small}>
              <a href={`${DOCS}/docs/collaboration`} className={bodyLink}>Collaboration guide</a>
            </p>
          </SectionHead>
          <div className="mt-[clamp(3rem,6vw,4.5rem)] grid gap-px border border-line-soft bg-line-soft lg:grid-cols-12">
            <div className="bg-surface p-[clamp(1rem,3vw,2.5rem)] pb-0 text-fg lg:col-span-7 [&_svg]:-mb-px">
              <CoEdit />
            </div>
            <ul className="grid list-none gap-px lg:col-span-5">
              {PEERS.map((peer) => (
                <li className="flex flex-col gap-3 bg-bg p-7 max-sm:p-6" key={peer.name}>
                  <h3 className={itemName}>{peer.name}</h3>
                  <p className={small}>{peer.desc}</p>
                </li>
              ))}
            </ul>
          </div>
        </div>
      </section>

      <section aria-labelledby="packages">
        <div className={`wrap rule ${pad}`}>
          <SectionHead
            n={4 + BENCH}
            id="packages"
            name={PACKAGES_SECTION.label}
            heading={PACKAGES_SECTION.heading}
          >
            <p className={lead}>{PACKAGES_SECTION.prose}</p>
            <p className={small}>
              See the <a href={`${DOCS}/docs/packages`} className={bodyLink}>package guide</a> for
              supported runtimes and fonts, and <a href={RELEASES} className={bodyLink}>release notes</a> for
              version-specific changes.
            </p>
          </SectionHead>

          <div className="mt-[clamp(3rem,6vw,4.5rem)] border border-line bg-surface">
            {ECOSYSTEMS.map((eco, i) => (
              <div
                key={eco.name}
                className={`grid items-center gap-x-6 gap-y-2 px-5 py-4 md:grid-cols-[12rem_1fr] ${
                  i > 0 ? "border-t border-line-soft" : ""
                }`}
              >
                <p className="flex items-baseline gap-3 font-mono text-[0.6875rem]">
                  <span className="uppercase tracking-[0.14em] text-fg">{eco.name}</span>
                  <a href={eco.url} target="_blank" rel="noopener" className={miniLink}>
                    {eco.registry} <ArrowUpRight size={11} strokeWidth={2} />
                  </a>
                  <a href={eco.docs} className={miniLink}>
                    docs <ArrowUpRight size={11} strokeWidth={2} />
                  </a>
                </p>
                <div className="flex min-w-0 items-center gap-3 font-mono text-[0.8125rem]">
                  <span className="text-dim select-none">$</span>
                  <code className="min-w-0 flex-1 py-1 [overflow-wrap:anywhere] sm:overflow-x-auto sm:whitespace-nowrap sm:[scrollbar-width:none] sm:[&::-webkit-scrollbar]:hidden">
                    {eco.install}
                  </code>
                  <CopyButton text={eco.install} />
                </div>
              </div>
            ))}
          </div>

          <ul className="mt-px list-none border-x border-b border-line-soft">
            {PACKAGES.map((pkg, i) => (
              <li
                key={pkg.name}
                className={`grid gap-x-10 gap-y-2 px-5 py-6 md:grid-cols-[16rem_1fr_auto] ${
                  i > 0 ? "border-t border-line-soft" : ""
                }`}
              >
                <h3 className="font-mono text-[0.8125rem] text-fg">{pkg.name}</h3>
                <p className={`${small} max-w-[42rem]`}>{pkg.desc}</p>
                <span className="flex gap-4 md:flex-col md:items-end md:gap-1.5">
                  <a
                    href={`https://www.npmjs.com/package/${pkg.name}`}
                    target="_blank"
                    rel="noopener"
                    className={miniLink}
                  >
                    npm <ArrowUpRight size={11} strokeWidth={2} />
                  </a>
                  <a href={`${DOCS}/docs/javascript`} className={miniLink}>
                    guide <ArrowUpRight size={11} strokeWidth={2} />
                  </a>
                </span>
              </li>
            ))}
          </ul>
        </div>
      </section>

      <section aria-labelledby="coda" className="overflow-hidden">
        <div className={`wrap rule ${pad} grid items-end gap-x-12 gap-y-12 lg:grid-cols-12`}>
          <div className="lg:col-span-8">
            <h2 id="coda" className={`${h2} max-w-[14ch]`}>
              {CODA.heading}
            </h2>
            <div className="mt-10 flex flex-wrap gap-3">
              <a href={DEMO} className={btnPrimary}>
                Try the demo <ArrowRight size={14} strokeWidth={2} />
              </a>
              <a href={DOCS} className={btnGhost}>
                Documentation <ArrowUpRight size={14} strokeWidth={2} />
              </a>
              <a href={REPO} target="_blank" rel="noopener" className={btnGhost}>
                GitHub <ArrowUpRight size={14} strokeWidth={2} />
              </a>
            </div>
            <p className={`${small} mt-8`}>
              <a href={DEMO} className={bodyLink}>Try it out</a>, or{" "}
              <a href={DOCS} className={bodyLink}>explore the docs</a> for setup, APIs, and format support.
            </p>
          </div>
          <div className="flex justify-end max-lg:hidden lg:col-span-4" aria-hidden="true">
            <Logo height={200} className="opacity-[0.07]" />
          </div>
        </div>
      </section>
    </>
  );
}
