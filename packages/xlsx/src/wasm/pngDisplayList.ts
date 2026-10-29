import type { Viewport, WorkbookHandle } from './loader';

export const pngDisplayListReaders = new WeakMap<WorkbookHandle, {
  viewport(viewport: Viewport): string;
  range(options?: { range?: string; scale?: number }): string;
}>();

export function workbookPngDisplayListJson(handle: WorkbookHandle, viewport: Viewport): string {
  const read = pngDisplayListReaders.get(handle);
  if (!read) throw new Error('Workbook display list is unavailable');
  return read.viewport(viewport);
}

export function workbookRangePngDisplayListJson(
  handle: WorkbookHandle, options?: { range?: string; scale?: number }
): string {
  const read = pngDisplayListReaders.get(handle);
  if (!read) throw new Error('Workbook display list is unavailable');
  return read.range(options);
}
