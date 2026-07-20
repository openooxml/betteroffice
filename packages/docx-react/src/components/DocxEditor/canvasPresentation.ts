import type { DisplayList } from '@betteroffice/docx/layout/render';

export function displayListNeedsHostImages(displayList: DisplayList): boolean {
  const visit = (value: unknown): boolean => {
    if (!value || typeof value !== 'object') return false;
    if (Array.isArray(value)) return value.some(visit);
    const record = value as Record<string, unknown>;
    if (record.kind === 'image' || record.kind === 'picture') return true;
    return Object.values(record).some(visit);
  };
  return visit(displayList.pages);
}
