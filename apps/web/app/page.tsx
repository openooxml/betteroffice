import type { Metadata } from "next";
import { ArrowRight, ArrowUpRight } from "lucide-react";
import { Cmd } from "./components/Cmd";
import { HeroField } from "./components/HeroField";
import { FadeIn, Reveal } from "./components/motion";

export const metadata: Metadata = {
  alternates: { canonical: "/" },
};

const REPO = "https://github.com/openooxml/betteroffice";
const DOCS = "https://docs.betteroffice.dev";

const sec = "sec relative border-t border-line-soft px-8 py-18 max-[44rem]:px-5";
const secLabel =
  "mb-8 flex items-baseline gap-3 font-mono text-[0.6875rem] uppercase tracking-[0.16em] text-dim";
const secH2 = "mb-3.5 text-[1.375rem] leading-[1.3] font-semibold tracking-[-0.02em]";
const secP = "max-w-[36rem] text-ink";
const bodyLink =
  "underline decoration-faint underline-offset-[3px] transition-colors hover:decoration-fg";
const btn =
  "inline-flex items-center gap-2 rounded-md border px-4.5 py-2.5 font-mono text-[0.8125rem] no-underline transition-colors";
const grid =
  "grid grid-cols-2 gap-px overflow-hidden rounded-md border border-line-soft bg-line-soft max-[44rem]:grid-cols-1";
const card = "flex flex-col gap-2 bg-bg p-6 transition-colors hover:bg-surface";
const cardName =
  "flex items-center justify-between gap-3 font-mono text-[0.8125rem] text-fg";
const cardLink =
  "inline-flex items-center gap-1 font-mono text-[0.6875rem] text-dim no-underline hover:text-fg";
const demoLink =
  "ml-auto inline-flex items-center gap-1 rounded border border-line px-2 py-0.5 font-mono text-[0.6875rem] text-ink no-underline transition-colors hover:border-dim hover:text-fg";

const EDITORS = [
  {
    name: "Documents",
    format: "docx",
    desc: "Word-faithful editing: fonts, theme colors, styles, tables, headers & footers, tracked changes.",
    live: true,
  },
  {
    name: "Spreadsheets",
    format: "xlsx",
    desc: "Calculation graph, grid rendering and number formats on the same shared core.",
    live: true,
  },
  {
    name: "Slides",
    format: "pptx",
    desc: "Slide model, masters and shape editing on the same shared core.",
    live: true,
  },
];

const PACKAGES = [
  {
    name: "@betteroffice/docx",
    desc: "Framework-free .docx core — parsing, CRDT editing and page layout in Rust, compiled to WebAssembly.",
  },
  {
    name: "@betteroffice/docx-react",
    desc: "The full DOCX editor as a React component — toolbar, pages, comments, tracked changes.",
  },
  {
    name: "@betteroffice/xlsx",
    desc: "Framework-free spreadsheet core — parsing, calculation and rendering on the Rust engine.",
  },
  {
    name: "@betteroffice/xlsx-react",
    desc: "The spreadsheet editor as a drop-in React component.",
  },
  {
    name: "@betteroffice/pptx",
    desc: "Framework-free slides core — parsing, editing and rendering on the Rust engine.",
  },
  {
    name: "@betteroffice/pptx-react",
    desc: "The slides editor as a drop-in React component.",
  },
];

const CAPABILITIES = [
  {
    name: "Own engines",
    desc: "We build the OOXML engines ourselves, in Rust — from the file format up. No wrapper around someone else's suite.",
  },
  {
    name: "Native OOXML editing",
    desc: "Documents are edited in their own format. No lossy conversion on open, none on save.",
  },
  {
    name: "Word-faithful output",
    desc: "What you see is what Word shows — layout, pagination and styling match the original.",
  },
  {
    name: "Real-time collaboration",
    desc: "The document is a CRDT — concurrent edits merge in the engine, not on a server.",
  },
  {
    name: "Agent-ready",
    desc: "Runs headless too — parse, edit and render documents server-side or inside agent pipelines.",
  },
  {
    name: "Apache 2.0",
    desc: "Permissive license, developed in the open, self-hostable without exceptions.",
  },
];

const JSON_LD = {
  "@context": "https://schema.org",
  "@type": "SoftwareApplication",
  name: "BetterOffice",
  url: "https://betteroffice.dev",
  image: "https://betteroffice.dev/logo.svg",
  applicationCategory: "BusinessApplication",
  operatingSystem: "Web",
  description:
    "The open-source office suite by the OpenOOXML project. Word-faithful DOCX and XLSX editors for React and JavaScript with real-time collaboration, on native OOXML engines written in Rust.",
  offers: {
    "@type": "Offer",
    price: "0",
    priceCurrency: "USD",
  },
  license: "https://www.apache.org/licenses/LICENSE-2.0",
  creator: {
    "@type": "Organization",
    name: "OpenOOXML",
    url: "https://openooxml.org",
  },
  sameAs: [
    "https://github.com/openooxml/betteroffice",
    "https://www.npmjs.com/org/betteroffice",
    "https://openooxml.org",
  ],
};

export default function Home() {
  return (
    <>
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(JSON_LD) }}
      />

      <section className="relative overflow-hidden px-8 pt-26 pb-18 max-[44rem]:px-5">
        <HeroField />
        <FadeIn delay={0.08} className="relative z-1">
          <h1 className="mb-4 text-[clamp(2.5rem,8vw,3.75rem)] leading-[1.05] font-[650] tracking-[-0.035em]">
            BetterOffice
          </h1>
        </FadeIn>
        <FadeIn delay={0.16} className="relative z-1">
          <p className="mb-10 max-w-[30rem] text-lg text-ink">
            The open-source office suite. Word-faithful editing and
            real-time collaboration on engines we build ourselves — running
            entirely in your browser, by the OpenOOXML project.
          </p>
        </FadeIn>
        <FadeIn delay={0.24} className="relative z-1">
          <div className="flex flex-wrap gap-3">
            <a
              href="https://demo.betteroffice.dev"
              className={`${btn} border-fg bg-fg text-bg hover:border-[#333333] hover:bg-[#333333]`}
            >
              Try the demo <ArrowRight size={14} strokeWidth={2} />
            </a>
            <a
              href={DOCS}
              className={`${btn} border-line text-ink hover:border-dim hover:text-fg`}
            >
              Documentation <ArrowUpRight size={14} strokeWidth={2} />
            </a>
            <a
              href={REPO}
              target="_blank"
              rel="noopener"
              className={`${btn} border-line text-ink hover:border-dim hover:text-fg`}
            >
              GitHub <ArrowUpRight size={14} strokeWidth={2} />
            </a>
          </div>
        </FadeIn>
      </section>

      <section className={sec} aria-labelledby="suite">
        <Reveal>
          <p className={secLabel}>
            <span className="text-faint">01</span> Suite
          </p>
          <h2 id="suite" className={secH2}>
            One suite, three editors
          </h2>
          <p className={`${secP} mb-6`}>
            BetterOffice packages the OpenOOXML engines as ready-to-use
            editors. Documents, spreadsheets and slides are all live today on
            the same foundation.
          </p>
        </Reveal>
        <Reveal delay={0.1}>
          <div className={grid}>
            {EDITORS.map((editor) => (
              <article className={card} key={editor.format}>
                <span className={cardName}>
                  {editor.name}
                  <span
                    className={`inline-flex shrink-0 items-center gap-1.5 font-mono text-[0.625rem] uppercase tracking-[0.12em] whitespace-nowrap ${
                      editor.live ? "text-acc" : "text-dim"
                    }`}
                  >
                    <span
                      className={`size-1.5 rounded-full ${
                        editor.live
                          ? "bg-acc shadow-[0_0_8px_rgba(5,150,105,0.45)]"
                          : "bg-faint"
                      }`}
                    />
                    {editor.live ? "available" : "coming"}
                  </span>
                </span>
                <span className="text-[0.8125rem] text-ink">{editor.desc}</span>
                <span className="mt-auto flex gap-4 pt-3">
                  <span className="rounded border border-line-soft bg-surface px-1.5 py-0.5 font-mono text-[0.7rem] whitespace-nowrap">
                    .{editor.format}
                  </span>
                  <a
                    href={`https://demo.betteroffice.dev/${editor.format}`}
                    className={demoLink}
                  >
                    demo <ArrowUpRight size={11} strokeWidth={2} />
                  </a>
                </span>
              </article>
            ))}
          </div>
        </Reveal>
      </section>

      <section className={sec} aria-labelledby="packages">
        <Reveal>
          <p className={secLabel}>
            <span className="text-faint">02</span> Packages
          </p>
          <h2 id="packages" className={secH2}>
            Ships as components, not iframes
          </h2>
          <p className={`${secP} mb-6`}>
            The editors install from npm and render inside your app — no
            embeds, no external services, documents never leave the page.
          </p>
        </Reveal>
        <Reveal delay={0.1}>
          <div className={`${grid} mb-6`}>
            {PACKAGES.map((pkg) => (
              <article className={card} key={pkg.name}>
                <span className={cardName}>{pkg.name}</span>
                <span className="text-[0.8125rem] text-ink">{pkg.desc}</span>
                <span className="mt-auto flex gap-4 pt-3">
                  <a
                    href={`https://www.npmjs.com/package/${pkg.name}`}
                    target="_blank"
                    rel="noopener"
                    className={cardLink}
                  >
                    npm <ArrowUpRight size={11} strokeWidth={2} />
                  </a>
                  <a
                    href={REPO}
                    target="_blank"
                    rel="noopener"
                    className={cardLink}
                  >
                    github <ArrowUpRight size={11} strokeWidth={2} />
                  </a>
                </span>
              </article>
            ))}
          </div>
          <Cmd command="npm install @betteroffice/docx-react" />
        </Reveal>
      </section>

      <section className={sec} aria-labelledby="foundation">
        <Reveal>
          <p className={secLabel}>
            <span className="text-faint">03</span> Foundation
          </p>
          <h2 id="foundation" className={secH2}>
            Built on our own engines
          </h2>
          <p className={secP}>
            BetterOffice is built by{" "}
            <a
              href="https://openooxml.org"
              target="_blank"
              rel="noopener"
              className={bodyLink}
            >
              OpenOOXML
            </a>
            , the open-source project writing native OOXML engines in Rust —
            parsing, layout, editing and rendering, from the file format up.
            Owning the whole stack is what makes the output Word-faithful.
          </p>
        </Reveal>
        <Reveal delay={0.1}>
          <ul className={`${grid} mt-8 list-none`}>
            {CAPABILITIES.map((cap) => (
              <li className="bg-bg px-5 py-4.5" key={cap.name}>
                <span className="mb-1 block font-mono text-xs text-fg">
                  {cap.name}
                </span>
                <span className="text-[0.8125rem] text-ink">{cap.desc}</span>
              </li>
            ))}
          </ul>
        </Reveal>
      </section>

      <section className={sec} aria-labelledby="collaboration">
        <Reveal>
          <p className={secLabel}>
            <span className="text-faint">04</span> Collaboration
          </p>
          <h2 id="collaboration" className={secH2}>
            People and agents, one document
          </h2>
          <p className={`${secP} mb-6`}>
            The document itself is a CRDT: every editor — every person, every
            AI agent — is a peer on the same data structure, and concurrent
            edits merge in the engine. Agents don&apos;t get a sidebar; they
            get a cursor, with the same undo and the same tracked-changes
            attribution as any co-author.
          </p>
        </Reveal>
        <Reveal delay={0.1}>
          <ul className={`${grid} list-none`}>
            <li className="bg-bg px-5 py-4.5">
              <span className="mb-1 block font-mono text-xs text-fg">
                People
              </span>
              <span className="text-[0.8125rem] text-ink">
                Live co-editing over any WebSocket relay. Offline edits
                converge on reconnect — merging is the data structure, not a
                server feature.
              </span>
            </li>
            <li className="bg-bg px-5 py-4.5">
              <span className="mb-1 block font-mono text-xs text-fg">
                Agents
              </span>
              <span className="text-[0.8125rem] text-ink">
                An agent edits through the same operations as a person, and
                human review is suggesting mode — accept or reject tracked
                changes, not a diff dialog.
              </span>
            </li>
          </ul>
        </Reveal>
      </section>
    </>
  );
}
