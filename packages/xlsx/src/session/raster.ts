import { rasterWasmAssetUrl } from '../wasm/asset';

type Raster = typeof import('../wasm/engineRaster');
let initialization: Promise<Raster> | undefined;

/** Longest wait for the next rasterizer byte before the export fails. */
export const RASTER_STALL_TIMEOUT_MS = 30_000;
/** Backstop for the whole rasterizer load, however steadily bytes arrive. */
export const RASTER_LOAD_TIMEOUT_MS = 300_000;

export interface RasterLoadOptions {
  stallMs?: number;
  totalMs?: number;
  /** Called while rasterizer bytes arrive, so the session can report the export as alive. */
  progress?: () => void;
}

function loadTimeout(message: string): Error {
  return Object.assign(new Error(`PNG export failed: ${message}`), {
    name: 'WorkbookPngRasterLoadError', code: 'png-raster-timeout',
  });
}

function loadRaster(options: RasterLoadOptions): Promise<Raster> {
  const stallMs = options.stallMs ?? RASTER_STALL_TIMEOUT_MS;
  const totalMs = options.totalMs ?? RASTER_LOAD_TIMEOUT_MS;
  const abort = new AbortController();
  return new Promise<Raster>((resolve, reject) => {
    let done = false;
    let stall: ReturnType<typeof setTimeout> | undefined;
    const total = setTimeout(() => fail(loadTimeout(`the rasterizer did not load within ${totalMs} ms`)), totalMs);
    function settle() {
      done = true;
      clearTimeout(stall);
      clearTimeout(total);
    }
    function fail(error: unknown) {
      if (done) return;
      settle();
      abort.abort();
      reject(error);
    }
    function alive() {
      if (done) return;
      clearTimeout(stall);
      stall = setTimeout(() => fail(loadTimeout(`no rasterizer bytes arrived for ${stallMs} ms`)), stallMs);
    }
    alive();
    (async () => {
      const raster = await import('../wasm/engineRaster');
      alive();
      const response = await fetch(rasterWasmAssetUrl(), { signal: abort.signal });
      alive();
      const body = response.body ? new Response(response.body.pipeThrough(new TransformStream<Uint8Array, Uint8Array>({
        transform(chunk, controller) {
          alive();
          try { options.progress?.(); } catch {}
          controller.enqueue(chunk);
        },
      })), { status: response.status, statusText: response.statusText, headers: response.headers }) : response;
      await raster.default({ module_or_path: body });
      return raster;
    })().then((raster) => {
      if (done) return;
      settle();
      resolve(raster);
    }, fail);
  });
}

export async function renderPng(displayListJson: string, options: RasterLoadOptions = {}): Promise<Uint8Array> {
  const pending = initialization ??= loadRaster(options);
  let raster: Raster;
  try { raster = await pending; }
  catch (error) {
    if (initialization === pending) initialization = undefined;
    throw error;
  }
  return raster.render_png(displayListJson);
}
