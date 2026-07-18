/**
 * Framework-agnostic print helpers shared by the React and Vue
 * adapters. Lifted from packages/docx-react/src/components/ui/PrintPreview.tsx
 * so both adapters use the same parsing / preview-window code path.
 *
 * The thin button component + the print-time CSS injection stay
 * adapter-local (they're framework-specific JSX/SFC bits); the data
 * helpers below are pure functions.
 */

export interface PrintOptions {
  includeHeaders?: boolean;
  includeFooters?: boolean;
  includePageNumbers?: boolean;
  pageRange?: { start: number; end: number } | null;
  scale?: number;
  printBackground?: boolean;
  margins?: 'default' | 'none' | 'minimum';
}

const DEFAULT_PRINT_OPTIONS: PrintOptions = {
  includeHeaders: true,
  includeFooters: true,
  includePageNumbers: true,
  pageRange: null,
  scale: 1.0,
  printBackground: true,
  margins: 'default',
};

export function getDefaultPrintOptions(): PrintOptions {
  return { ...DEFAULT_PRINT_OPTIONS };
}

/** Trigger browser print dialog for the current document. */
export function triggerPrint(): void {
  if (typeof window !== 'undefined') window.print();
}

const PRINT_CSS = '@media print { body { margin: 0; padding: 0; } @page { margin: 0; } }';

/**
 * Open a new window with print-optimised body content.
 *
 * Built entirely via DOM APIs (no `document.write` of interpolated strings):
 * `title` is assigned as a property so a value like `</title><script>` is
 * treated as text and cannot break out, and `content` is parsed in an inert
 * document and imported rather than concatenated into markup. `content` is the
 * caller's already-rendered print HTML; provide trusted markup.
 * Sibling copy: packages/docx-react/src/components/ui/PrintPreview.tsx.
 */
export function openPrintWindow(title: string = 'Document', content: string): Window | null {
  if (typeof window === 'undefined') return null;
  const w = window.open('', '_blank');
  if (!w) return null;
  // Sever the popup's back-reference to the opener (reverse-tabnabbing defence).
  w.opener = null;
  const doc = w.document;

  doc.title = title;

  const style = doc.createElement('style');
  style.textContent = PRINT_CSS;
  doc.head.appendChild(style);

  const parsed = new DOMParser().parseFromString(content, 'text/html');
  for (const node of Array.from(parsed.body.childNodes)) {
    doc.body.appendChild(doc.importNode(node, true));
  }

  return w;
}

/** Parse "1", "1-5", etc. into a page range, or null on invalid. */
export function parsePageRange(
  input: string,
  maxPages: number
): { start: number; end: number } | null {
  if (!input || !input.trim()) return null;
  const t = input.trim();
  if (/^\d+$/.test(t)) {
    const p = parseInt(t, 10);
    return p >= 1 && p <= maxPages ? { start: p, end: p } : null;
  }
  const m = t.match(/^(\d+)-(\d+)$/);
  if (m) {
    const start = parseInt(m[1], 10);
    const end = parseInt(m[2], 10);
    if (start >= 1 && end <= maxPages && start <= end) return { start, end };
  }
  return null;
}

export function formatPageRange(
  range: { start: number; end: number } | null,
  totalPages: number
): string {
  if (!range) return `All (${totalPages} pages)`;
  if (range.start === range.end) return `Page ${range.start}`;
  return `Pages ${range.start}-${range.end}`;
}

export function isPrintSupported(): boolean {
  return typeof window !== 'undefined' && typeof window.print === 'function';
}
