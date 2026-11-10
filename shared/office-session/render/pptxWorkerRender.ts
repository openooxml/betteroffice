import type { SessionTransport } from '../transport';
import { createBitmapOwnership } from './ownership';
import type { DeviceRect } from './types';

export interface PptxPaintFont {
  family: string;
  source: ArrayBuffer | string;
  descriptors: FontFaceDescriptors;
}

export type PptxRenderRequest =
  | { kind: 'view'; scale: number; dpr: number }
  | { kind: 'fonts'; faces: readonly PptxPaintFont[] }
  | { kind: 'resync' }
  | { kind: 'prerender'; slideId: string; token: number }
  | { kind: 'dispose' };

export type PptxRenderReply =
  | { kind: 'raster'; slideId: string; frameEpoch: number; factor: number; baseEpoch: number | null;
    content: number; baseContent: number | null;
    rect: DeviceRect; bitmap: ImageBitmap; width: number; height: number }
  | { kind: 'unchanged'; slideId: string; frameEpoch: number; baseEpoch: number; baseContent: number }
  | { kind: 'main-only'; slideId: string; frameEpoch: number; assetIds: string[] }
  | { kind: 'error'; frameEpoch: number; message: string }
  | { kind: 'prerender'; slideId: string; token: number; version: string; fonts: number; factor: number; width: number;
    height: number; bitmap: ImageBitmap }
  | { kind: 'font-generation'; fonts: number }
  | { kind: 'unsupported' };

export interface PptxWorkerRenderClient {
  readonly supported: boolean;
  view(scale: number, dpr: number): void;
  fonts(faces: readonly PptxPaintFont[]): void;
  resync(): void;
  /** Asks for an ahead-of-time raster of a slide at the current view; returns the token its reply carries. */
  prerender(slideId: string): number;
  idle(): void;
  subscribe(listener: (reply: PptxRenderReply) => void): () => void;
  dispose(): void;
}

const scope = globalThis as typeof globalThis & {
  __betterofficePptxWorkerRender?: { options?: { enabled?: boolean } };
  __betterofficePptxRasterOwnership?: ReturnType<typeof createBitmapOwnership<ImageBitmap>>;
  __betterofficePptxWorkerRenders?: WeakMap<object, PptxWorkerRenderClient>;
};
export const pptxWorkerRenderClients = scope.__betterofficePptxWorkerRenders ??= new WeakMap();
export const pptxRasterOwnership = scope.__betterofficePptxRasterOwnership ??= createBitmapOwnership<ImageBitmap>();
export const PPTX_RENDER_PORT = 'pptx-render-environment';
export const PPTX_RENDER_REPLY = 'pptxRender';

export function attachPptxWorkerRender(transport: SessionTransport): PptxWorkerRenderClient | undefined {
  if (scope.__betterofficePptxWorkerRender?.options?.enabled === false || typeof MessageChannel !== 'function') return;
  let channel: MessageChannel;
  try { channel = new MessageChannel(); } catch { return; }
  const { port1, port2 } = channel;
  const listeners = new Set<(reply: PptxRenderReply) => void>();
  const ownership = pptxRasterOwnership;
  let disposed = false;
  let supported = true;
  let offError = () => {};
  let offReplies = () => {};
  let sentView: { scale: number; dpr: number } | undefined;
  let prerenders = 0;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const clear = () => { if (timer !== undefined) clearTimeout(timer); timer = undefined; };
  const fail = () => {
    if (disposed || !supported) return;
    supported = false; clear();
    try { port1.postMessage({ kind: 'dispose' }); } catch {}
    for (const listener of [...listeners]) listener({ kind: 'unsupported' });
  };
  const send = (message: PptxRenderRequest, transfer: Transferable[] = []) => {
    if (disposed || !supported) return;
    try { port1.postMessage(message, transfer); } catch { fail(); }
  };
  const client: PptxWorkerRenderClient = {
    get supported() { return supported && !disposed; },
    view(scale, dpr) {
      if (timer === undefined && supported && !disposed) timer = setTimeout(fail, 5000);
      if (sentView?.scale === scale && sentView.dpr === dpr) return;
      sentView = { scale, dpr };
      send({ kind: 'view', scale, dpr });
    },
    fonts(faces) {
      const copies = faces.map((face) => ({ ...face, descriptors: { ...face.descriptors },
        source: typeof face.source === 'string' ? face.source : face.source.slice(0) }));
      send({ kind: 'fonts', faces: copies }, copies.flatMap((face) => typeof face.source === 'string' ? [] : [face.source]));
    },
    resync: () => send({ kind: 'resync' }),
    prerender(slideId) {
      const token = ++prerenders;
      send({ kind: 'prerender', slideId, token });
      return token;
    },
    idle: clear,
    subscribe(listener) {
      if (!client.supported) listener({ kind: 'unsupported' });
      if (!disposed) listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    dispose() {
      if (disposed) return;
      fail();
      disposed = true;
      offError(); offReplies(); listeners.clear(); port1.close();
    },
  };
  const receive = (reply: PptxRenderReply) => {
    if (reply.kind !== 'font-generation' && reply.kind !== 'prerender') clear();
    if (reply.kind === 'unsupported') { fail(); return; }
    const owned = reply.kind === 'raster' || reply.kind === 'prerender' ? reply.bitmap : undefined;
    if (owned) ownership.retain(owned, reply);
    try { if (!disposed) for (const listener of [...listeners]) listener(reply); }
    finally { if (owned) ownership.release(owned, reply); }
  };
  port1.onmessage = (event: MessageEvent<PptxRenderReply>) => receive(event.data);
  offReplies = transport.listen((message) => {
    if (!message || typeof message !== 'object' || !(PPTX_RENDER_REPLY in message)) return;
    const reply = (message as Record<string, unknown>)[PPTX_RENDER_REPLY] as PptxRenderReply;
    delete (message as Record<string, unknown>)[PPTX_RENDER_REPLY];
    if (disposed || !supported) { if (reply.kind === 'raster') reply.bitmap.close(); return; }
    receive(reply);
  });
  port1.onmessageerror = fail;
  port1.addEventListener('close', fail);
  offError = transport.onError(fail);
  try { transport.post({ kind: PPTX_RENDER_PORT, port: port2 }, [port2]); }
  catch { client.dispose(); port2.close(); return; }
  return client;
}
