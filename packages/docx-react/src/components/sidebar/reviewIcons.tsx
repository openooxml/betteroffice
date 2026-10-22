import type { ReactNode } from 'react';

// Path data from lucide-static@1.61.0 icons/check.svg and icons/x.svg (ISC; notices in the package THIRD-PARTY-NOTICES.md).
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
