import type { ReactNode } from 'react';

// Path data from lucide-static@1.61.0 icons/check.svg, x.svg, undo-2.svg, ellipsis-vertical.svg and message-square-check.svg (ISC; notices in the package THIRD-PARTY-NOTICES.md).
function LucideIcon({ children }: { children: ReactNode }) {
  return <svg className="docx-review-icon" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none"
    stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{children}</svg>;
}

export function ReviewCheckIcon() {
  return <LucideIcon><path d="M20 6 9 17l-5-5" /></LucideIcon>;
}

export function ReviewXIcon() {
  return <LucideIcon><path d="M18 6 6 18" /><path d="m6 6 12 12" /></LucideIcon>;
}

export function ReviewUndoIcon() {
  return <LucideIcon><path d="M9 14 4 9l5-5" /><path d="M4 9h10.5a5.5 5.5 0 0 1 5.5 5.5a5.5 5.5 0 0 1-5.5 5.5H11" /></LucideIcon>;
}

export function ReviewMoreIcon() {
  return <LucideIcon><circle cx="12" cy="12" r="1" /><circle cx="12" cy="5" r="1" /><circle cx="12" cy="19" r="1" /></LucideIcon>;
}

export function ReviewResolvedIcon() {
  return <LucideIcon><path d="M22 17a2 2 0 0 1-2 2H6.828a2 2 0 0 0-1.414.586l-2.202 2.202A.7.7 0 0 1 2 21.286V5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2z" /><path d="m9 11 2 2 4-4" /></LucideIcon>;
}
