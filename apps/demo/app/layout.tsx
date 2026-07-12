import type { Metadata, Viewport } from "next";
import Link from "next/link";
import { Logo } from "./components/Logo";
import { formats } from "../lib/formats";
import "./globals.css";

const SITE = "https://demo.betteroffice.dev";

export const metadata: Metadata = {
  metadataBase: new URL(SITE),
  title: {
    default: "BetterOffice — Demos",
    template: "%s — BetterOffice Demos",
  },
  description:
    "Live demos of the BetterOffice engines — Word documents, spreadsheets, and slides, rendered by native OOXML engines in Rust.",
};

export const viewport: Viewport = {
  themeColor: "#ffffff",
};

const navLink =
  "font-mono text-xs text-ink no-underline transition-colors hover:text-fg";
const footLink = "font-mono text-xs text-ink no-underline hover:text-fg";

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>
        <div className="mx-auto flex min-h-dvh max-w-[44rem] flex-col border-x border-line-soft max-[44rem]:border-x-0">
          <header className="sticky top-0 z-10 flex items-center justify-between border-b border-line-soft bg-white/80 px-8 py-4 backdrop-blur-xl max-[44rem]:flex-wrap max-[44rem]:gap-y-2 max-[44rem]:px-5 max-[44rem]:py-3.5">
            <Link
              href="/"
              className="flex items-center gap-2.5 text-[0.9375rem] font-semibold tracking-[-0.01em] no-underline"
            >
              <Logo height={19} />
              BetterOffice{" "}
              <span className="font-normal text-faint">/ demos</span>
            </Link>
            <nav
              aria-label="Demo formats"
              className="flex items-center gap-5 max-[44rem]:w-full max-[44rem]:justify-between"
            >
              {formats.map((format) => (
                <Link
                  key={format.id}
                  href={`/${format.id}`}
                  className={navLink}
                >
                  {format.id}
                </Link>
              ))}
              <a href="https://betteroffice.dev" className={navLink}>
                website
              </a>
            </nav>
          </header>
          <main className="flex-1">{children}</main>
          <footer className="flex flex-wrap items-end justify-between gap-4 border-t border-line-soft px-8 py-8 max-[44rem]:px-5">
            <div className="flex gap-6">
              <a href="https://betteroffice.dev" className={footLink}>
                website
              </a>
              <a
                href="https://github.com/openooxml/betteroffice"
                className={footLink}
              >
                github
              </a>
            </div>
            <p className="font-mono text-[0.6875rem] text-dim">Apache 2.0</p>
          </footer>
        </div>
      </body>
    </html>
  );
}
