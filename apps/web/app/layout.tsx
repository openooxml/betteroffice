import type { Metadata, Viewport } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import Link from "next/link";
import { Logo } from "./components/Logo";
import { BENCHMARKS, BENCHMARKS_SITE, EVALS_SITE, SITE, SITE_DESCRIPTION, SITE_TITLE } from "./content";
import "./globals.css";

const sans = Geist({ subsets: ["latin"], variable: "--font-geist", display: "swap" });
const mono = Geist_Mono({ subsets: ["latin"], variable: "--font-geist-mono", display: "swap" });

const navLink =
  "font-mono text-xs text-ink no-underline transition-colors hover:text-fg";
const footHead =
  "mb-4 font-mono text-[0.6875rem] font-normal uppercase tracking-[0.14em] text-dim";
const footLink = "font-mono text-xs text-ink no-underline transition-colors hover:text-fg";

export const metadata: Metadata = {
  metadataBase: new URL(SITE),
  title: {
    default: SITE_TITLE,
    template: "%s — BetterOffice",
  },
  description: SITE_DESCRIPTION,
  keywords: [
    "BetterOffice",
    "open-source office suite",
    "Apache-2.0 document editor",
    "open source docx editor",
    "DOCX editor",
    "DOCX viewer",
    "react docx viewer",
    "react docx editor",
    "javascript docx viewer",
    "javascript docx editor",
    "render docx in browser",
    "XLSX editor",
    "react spreadsheet component",
    "javascript spreadsheet library",
    "PPTX editor",
    "react pptx viewer",
    "Word-compatible editor",
    "Excel-compatible spreadsheet",
    "OOXML",
    "OpenOOXML",
    "real-time collaboration",
    "CRDT document editing",
    "headless document processing",
    "Rust",
    "WebAssembly",
  ],
  openGraph: {
    type: "website",
    siteName: "BetterOffice",
    url: SITE,
    title: SITE_TITLE,
    description: SITE_DESCRIPTION,
  },
  twitter: {
    card: "summary_large_image",
    title: SITE_TITLE,
    description: SITE_DESCRIPTION,
  },
  robots: {
    index: true,
    follow: true,
  },
};

export const viewport: Viewport = {
  themeColor: [
    { media: "(prefers-color-scheme: light)", color: "#ffffff" },
    { media: "(prefers-color-scheme: dark)", color: "#090909" },
  ],
  colorScheme: "light dark",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en" className={`${sans.variable} ${mono.variable}`}>
      <body className="flex min-h-dvh flex-col">
        <header className="sticky top-0 z-20 border-b border-line-soft bg-head backdrop-blur-xl">
          <div className="wrap flex h-16 items-center justify-between gap-6 max-[36rem]:h-auto max-[36rem]:flex-wrap max-[36rem]:gap-y-2 max-[36rem]:py-3.5">
            <Link
              href="/"
              className="flex items-center gap-2.5 text-[0.9375rem] font-semibold tracking-[-0.015em] no-underline"
            >
              <Logo height={18} />
              BetterOffice
            </Link>
            <nav
              aria-label="Site"
              className="flex flex-wrap items-center gap-x-7 gap-y-2 max-[44rem]:gap-x-5 max-[36rem]:gap-x-4"
            >
              <a href="https://demo.betteroffice.dev" className={navLink}>
                demos
              </a>
              <a href="https://docs.betteroffice.dev" className={navLink}>
                docs
              </a>
              <a href={BENCHMARKS_SITE} className={navLink}>
                benchmarks
              </a>
              <a href={EVALS_SITE} className={navLink}>
                evals
              </a>
              <a
                href="https://github.com/openooxml/betteroffice"
                target="_blank"
                rel="noopener"
                className={navLink}
              >
                github
              </a>
            </nav>
          </div>
        </header>
        <main className="flex-1">{children}</main>
        <footer>
          <div className="wrap rule flex flex-wrap items-end justify-between gap-10 py-14 max-[48rem]:py-10">
            <div className="flex gap-16 max-[48rem]:flex-wrap max-[48rem]:gap-10">
              <div>
                <h3 className={footHead}>Product</h3>
                <ul className="flex list-none flex-col gap-2">
                  <li>
                    <a href="https://docs.betteroffice.dev" className={footLink}>
                      Docs
                    </a>
                  </li>
                  <li>
                    <a
                      href="https://www.npmjs.com/org/betteroffice"
                      target="_blank"
                      rel="noopener"
                      className={footLink}
                    >
                      npm
                    </a>
                  </li>
                </ul>
              </div>
              <div>
                <h3 className={footHead}>Source</h3>
                <ul className="flex list-none flex-col gap-2">
                  <li>
                    <a
                      href="https://github.com/openooxml/betteroffice"
                      target="_blank"
                      rel="noopener"
                      className={footLink}
                    >
                      GitHub
                    </a>
                  </li>
                  <li>
                    <a href={BENCHMARKS} className={footLink}>
                      Fidelity results
                    </a>
                  </li>
                  <li>
                    <a href="/llms.txt" className={footLink}>
                      llms.txt
                    </a>
                  </li>
                </ul>
              </div>
              <div>
                <h3 className={footHead}>Project</h3>
                <ul className="flex list-none flex-col gap-2">
                  <li>
                    <a
                      href="https://openooxml.org"
                      target="_blank"
                      rel="noopener"
                      className={footLink}
                    >
                      OpenOOXML
                    </a>
                  </li>
                </ul>
              </div>
            </div>
            <p className="font-mono text-[0.6875rem] text-dim">
              Apache License 2.0 — by the OpenOOXML project
            </p>
          </div>
        </footer>
      </body>
    </html>
  );
}
