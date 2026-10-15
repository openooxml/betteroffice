import type { Viewport } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import Link from "next/link";
import { Logo } from "../components/Logo";
import { listedLiveFormats } from "../../lib/formats";

const sans = Geist({ subsets: ["latin"], variable: "--font-geist", display: "swap" });
const mono = Geist_Mono({ subsets: ["latin"], variable: "--font-geist-mono", display: "swap" });

export const viewport: Viewport = {
  themeColor: [
    { media: "(prefers-color-scheme: light)", color: "#ffffff" },
    { media: "(prefers-color-scheme: dark)", color: "#090909" },
  ],
  colorScheme: "light dark",
};

const navLink = "font-mono text-xs text-ink no-underline transition-colors hover:text-fg";

// the launcher's chrome, in the website's frame and type; the editor routes render without it
export default function SiteLayout({ children }: { children: React.ReactNode }) {
  return (
    <div className={`site ${sans.variable} ${mono.variable} flex min-h-dvh flex-col antialiased`}>
      <header className="sticky top-0 z-10 border-b border-line-soft bg-head backdrop-blur-xl">
        <div className="wrap flex h-16 items-center justify-between gap-6 max-[36rem]:h-auto max-[36rem]:flex-wrap max-[36rem]:gap-y-2 max-[36rem]:py-3.5">
          <Link
            href="/"
            className="flex items-center gap-2.5 text-[0.9375rem] font-semibold tracking-[-0.015em] text-fg no-underline"
          >
            <Logo height={18} />
            BetterOffice <span className="font-normal text-faint">/ demos</span>
          </Link>
          <nav aria-label="Demo formats" className="flex flex-wrap items-center gap-x-7 gap-y-2 max-[44rem]:gap-x-5">
            {listedLiveFormats.map((format) => (
              <Link key={format.id} href={`/${format.id}`} className={navLink}>
                {format.id}
              </Link>
            ))}
            <a href="https://betteroffice.dev" className={navLink}>
              website
            </a>
          </nav>
        </div>
      </header>
      <main className="flex-1">{children}</main>
      <footer>
        <div className="wrap rule flex flex-wrap items-end justify-between gap-6 py-10">
          <div className="flex gap-6">
            <a href="https://betteroffice.dev" className={navLink}>
              website
            </a>
            <a href="https://github.com/openooxml/betteroffice" className={navLink}>
              github
            </a>
          </div>
          <p className="font-mono text-[0.6875rem] text-dim">Apache 2.0</p>
        </div>
      </footer>
    </div>
  );
}
