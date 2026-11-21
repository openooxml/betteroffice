import { createFontPackResponder, type FontPackReply, type OptionalFontPackSource } from './optionalFontPacks';

const importOptionalFontPack: OptionalFontPackSource = async (pack) =>
  (await import('@betteroffice/fonts')).optionalFontPackUrls(pack);
const ATTACHED = Symbol.for('betteroffice.fontPackSource');

/** Whether a font-pack responder is attached to `worker`, by any copy of this module. */
export function hasFontPackSource(worker: object): boolean {
  return (worker as { [ATTACHED]?: boolean })[ATTACHED] === true;
}

/** A responder for requests that reach a handler without an attached responder. */
export function createPageFontPackResponder(post: (message: FontPackReply) => void) {
  return createFontPackResponder(post, importOptionalFontPack);
}

export function attachFontPackSource(
  worker: Pick<Worker, 'postMessage' | 'addEventListener' | 'removeEventListener'>,
  source = importOptionalFontPack,
): (() => void) | undefined {
  if (typeof worker.addEventListener !== 'function' || hasFontPackSource(worker)) return;
  const responder = createFontPackResponder((message) => worker.postMessage(message), source);
  const receive = (event: MessageEvent) => {
    if (responder.receive(event.data)) event.stopImmediatePropagation();
  };
  worker.addEventListener('message', receive);
  (worker as { [ATTACHED]?: boolean })[ATTACHED] = true;
  return () => {
    worker.removeEventListener?.('message', receive);
    delete (worker as { [ATTACHED]?: boolean })[ATTACHED];
  };
}
