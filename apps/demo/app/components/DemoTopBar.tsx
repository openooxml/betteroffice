import Link from "next/link";
import type { ReactNode } from "react";
import { Logo } from "./Logo";

const control =
  "inline-flex h-7 flex-none cursor-pointer items-center justify-center rounded-[6px] px-2.5 text-[12.5px] leading-none font-medium whitespace-nowrap transition-colors duration-[140ms] ease-[ease] focus-visible:outline-2 focus-visible:outline-offset-1 focus-visible:outline-[#1f1f1f] disabled:cursor-default disabled:opacity-[0.38]";

/** Borderless text button; with `aria-pressed` it doubles as a toggle. */
export const barText = `${control} text-[#444746] hover:bg-[rgba(31,31,31,0.08)] hover:text-[#1f1f1f] disabled:bg-transparent disabled:text-[#444746] aria-pressed:bg-[#1f1f1f] aria-pressed:text-white`;
export const barBordered = `${control} border border-[rgba(31,31,31,0.16)] bg-white text-[#1f1f1f] hover:bg-[rgba(31,31,31,0.08)] disabled:bg-white`;
export const barPrimary = `${control} bg-[#1f1f1f] text-white hover:bg-[#3c3c3c] disabled:bg-[#1f1f1f]`;

export function BarDivider() {
  return (
    <span
      aria-hidden="true"
      className="mx-1 h-4 w-px flex-none bg-[rgba(31,31,31,0.16)]"
    />
  );
}

interface DemoTopBarProps {
  format: string;
  tagline: string;
  fileName?: string | null;
  /** The page's controls, in order; the GitHub link is appended. */
  children: ReactNode;
}

// Wide: one row, wrapping the controls below the brand when they do not fit.
// <= 720px: brand row, then one horizontally scrolling controls row.
export function DemoTopBar({
  format,
  tagline,
  fileName,
  children,
}: DemoTopBarProps) {
  return (
    <header className="z-2 flex flex-wrap items-center gap-x-3 gap-y-2 border-b border-hairline bg-white/92 px-4 py-2.5 backdrop-blur-lg">
      <Link
        href="/"
        className="inline-flex flex-none items-baseline gap-2 text-[14px] font-[650] tracking-[-0.01em] whitespace-nowrap text-fg no-underline"
      >
        <Logo height={18} className="self-center" />
        BetterOffice <span className="font-normal text-faint">/ {format}</span>
      </Link>
      <span className="min-w-0 flex-1 truncate text-[12.5px] text-mute max-[720px]:hidden">
        {tagline}
      </span>
      {fileName && (
        <span className="min-w-0 max-w-[180px] truncate text-[12.5px] text-mute max-[720px]:hidden">
          {fileName}
        </span>
      )}
      <div className="ml-auto flex max-w-full min-w-0 flex-wrap items-center justify-end gap-1.5 [scrollbar-width:none] max-[720px]:-my-1 max-[720px]:-mr-1 max-[720px]:-ml-1 max-[720px]:max-w-none max-[720px]:basis-[calc(100%+0.5rem)] max-[720px]:flex-nowrap max-[720px]:justify-start max-[720px]:overflow-x-auto max-[720px]:p-1 [&::-webkit-scrollbar]:hidden">
        {children}
        <a
          className="inline-flex size-7 flex-none items-center justify-center rounded-[6px] text-[#444746] transition-colors duration-[140ms] ease-[ease] hover:bg-[rgba(31,31,31,0.08)] hover:text-[#1f1f1f] focus-visible:outline-2 focus-visible:outline-offset-1 focus-visible:outline-[#1f1f1f]"
          href="https://github.com/openooxml/betteroffice"
          target="_blank"
          rel="noreferrer"
          aria-label="View on GitHub"
          title="View on GitHub"
        >
          <svg
            width="16"
            height="16"
            viewBox="0 0 16 16"
            fill="currentColor"
            aria-hidden="true"
          >
            <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z" />
          </svg>
        </a>
      </div>
    </header>
  );
}
