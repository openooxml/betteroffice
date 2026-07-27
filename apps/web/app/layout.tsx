import type { Metadata, Viewport } from "next";
import Link from "next/link";
import { Logo } from "./components/Logo";
import "./globals.css";

const SITE = "https://betteroffice.dev";

const navLink =
  "font-mono text-xs text-ink no-underline transition-colors hover:text-fg";
const footHead =
  "mb-3 font-mono text-[0.625rem] font-normal uppercase tracking-[0.16em] text-faint";
const footLink = "font-mono text-xs text-ink no-underline hover:text-fg";

export const metadata: Metadata = {
  metadataBase: new URL(SITE),
  title: {
    default: "BetterOffice — The open-source office suite",
    template: "%s — BetterOffice",
  },
  description:
    "Apache-2.0 DOCX, XLSX and PPTX editors and viewers for React and plain JavaScript, by the OpenOOXML project. Native OOXML engines written in Rust, compiled to WebAssembly, with real-time collaboration in the engine.",
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
    "eigenpal alternative",
  ],
  openGraph: {
    type: "website",
    siteName: "BetterOffice",
    url: SITE,
  },
  twitter: {
    card: "summary_large_image",
  },
  robots: {
    index: true,
    follow: true,
  },
};

export const viewport: Viewport = {
  themeColor: "#ffffff",
  colorScheme: "light",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>
        <div className="mx-auto flex min-h-dvh max-w-[44rem] flex-col border-x border-line-soft max-[44rem]:border-x-0">
          <header className="sticky top-0 z-10 flex items-center justify-between border-b border-line-soft bg-white/80 px-8 py-4 backdrop-blur-xl max-[44rem]:flex-wrap max-[44rem]:gap-y-1.5 max-[44rem]:px-5 max-[44rem]:py-3.5">
            <Link
              href="/"
              className="flex items-center gap-2.5 text-[0.9375rem] font-semibold tracking-[-0.01em] no-underline"
            >
              <Logo height={19} />
              BetterOffice
            </Link>
            <nav
              aria-label="Site"
              className="flex items-center gap-6 max-[44rem]:flex-wrap max-[44rem]:justify-end max-[44rem]:gap-4"
            >
              <a href="https://demo.betteroffice.dev" className={navLink}>
                demos
              </a>
              <a href="https://docs.betteroffice.dev" className={navLink}>
                docs
              </a>
              <a
                href="https://github.com/openooxml/betteroffice"
                target="_blank"
                rel="noopener"
                className={navLink}
              >
                github
              </a>
              <a
                href="https://openooxml.org"
                target="_blank"
                rel="noopener"
                className={navLink}
              >
                openooxml.org
              </a>
            </nav>
          </header>
          <main className="flex-1">{children}</main>
          <footer className="flex flex-wrap items-end justify-between gap-8 border-t border-line-soft px-8 py-10 max-[44rem]:px-5 max-[44rem]:py-8">
            <div className="flex gap-12 max-[44rem]:flex-wrap max-[44rem]:gap-8">
              <div>
                <h3 className={footHead}>Product</h3>
                <ul className="flex list-none flex-col gap-1.5">
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
                <ul className="flex list-none flex-col gap-1.5">
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
                    <a href="/llms.txt" className={footLink}>
                      llms.txt
                    </a>
                  </li>
                </ul>
              </div>
              <div>
                <h3 className={footHead}>Project</h3>
                <ul className="flex list-none flex-col gap-1.5">
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
                  <li>
                    <a
                      href="https://openooxml.org/eigenpal"
                      target="_blank"
                      rel="noopener"
                      className={footLink}
                    >
                      eigenpal migration
                    </a>
                  </li>
                </ul>
              </div>
            </div>
            <p className="font-mono text-[0.6875rem] text-dim">
              Apache License 2.0 — by the OpenOOXML project
            </p>
          </footer>
        </div>
      </body>
    </html>
  );
}
