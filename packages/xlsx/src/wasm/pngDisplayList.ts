import type { WorkbookHandle } from './loader';

/**
 * A session's lazily loaded rasterizer, for runtimes with neither a canvas nor
 * a rasterizer in the engine module. It draws a PNG display list.
 */
export const pngRasterizers = new WeakMap<WorkbookHandle, (displayListJson: string) => Promise<Uint8Array>>();
