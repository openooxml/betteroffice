import type { Layout } from '@betteroffice/docx/layout/pagination';

/** The document's page count, or 0 while `layout` covers only part of it. */
export function documentPageCount(layout: Layout | null | undefined): number {
  return layout && !layout.partial ? layout.pages.length : 0;
}
