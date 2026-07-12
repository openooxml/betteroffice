import type { Metadata, Viewport } from "next";
import Link from "next/link";
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

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body className="min-h-dvh bg-bg text-fg antialiased">
        <header className="flex items-center justify-between border-b border-line px-6 py-4">
          <Link href="/" className="font-mono text-sm text-fg no-underline">
            BetterOffice{" "}
            <span className="text-faint">/ demos</span>
          </Link>
          <nav className="flex items-center gap-5">
            {formats.map((f) => (
              <Link key={f.id} href={`/${f.id}`} className={navLink}>
                {f.name}
              </Link>
            ))}
            <a
              href="https://betteroffice.dev"
              className={navLink}
              target="_blank"
              rel="noreferrer"
            >
              betteroffice.dev ↗
            </a>
          </nav>
        </header>
        <main className="mx-auto max-w-5xl px-6 py-14">{children}</main>
      </body>
    </html>
  );
}
