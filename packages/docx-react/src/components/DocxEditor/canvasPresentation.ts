import type { DisplayList } from '@betteroffice/docx/layout/render';

// Cached per page object. Safe because frame deltas replace changed pages with
// new objects; in-place owned patches only move positions and cannot add or
// remove image content.
const pageNeedsHostImages = new WeakMap<object, boolean>();

function visit(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  if (Array.isArray(value)) return value.some(visit);
  const record = value as Record<string, unknown>;
  if (record.kind === 'image' || record.kind === 'picture') return true;
  return Object.values(record).some(visit);
}

export function displayListNeedsHostImages(displayList: DisplayList): boolean {
  for (const page of displayList.pages) {
    let needs = pageNeedsHostImages.get(page);
    if (needs === undefined) {
      needs = visit(page);
      pageNeedsHostImages.set(page, needs);
    }
    if (needs) return true;
  }
  return false;
}
