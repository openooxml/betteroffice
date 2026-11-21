import { createFontPackResponder, type FontPackReply, type OptionalFontPackSource } from './optionalFontPacks';

const importOptionalFontPack: OptionalFontPackSource = async (pack) =>
  (await import('@betteroffice/fonts')).optionalFontPackUrls(pack);
const ATTACHED = Symbol.for('betteroffice.fontPackSource');
const DETACH = Symbol.for('betteroffice.fontPackSource.detach');
type FontPackSourceOwner = { [ATTACHED]?: true; [DETACH]?: () => void };

/** Whether a font-pack responder is attached to `worker`, by any copy of this module. */
export function hasFontPackSource(worker: object): boolean {
  return (worker as FontPackSourceOwner)[ATTACHED] === true;
}

export function getFontPackSourceDetach(worker: object): (() => void) | undefined {
  return (worker as FontPackSourceOwner)[DETACH];
}

/** A responder for requests that reach a handler without an attached responder. */
export function createPageFontPackResponder(post: (message: FontPackReply) => void) {
  return createFontPackResponder(post, importOptionalFontPack);
}

export function attachFontPackSource(
  worker: Pick<Worker, 'postMessage' | 'addEventListener' | 'removeEventListener'>,
  source = importOptionalFontPack,
): (() => void) | undefined {
  const owner = worker as FontPackSourceOwner;
  if (owner[ATTACHED]) return owner[DETACH];
  if (typeof worker.addEventListener !== 'function') return;
  const responder = createFontPackResponder((message) => worker.postMessage(message), source);
  const receive = (event: MessageEvent) => {
    if (responder.receive(event.data)) event.stopImmediatePropagation();
  };
  const detach = () => {
    if (owner[DETACH] !== detach) return;
    worker.removeEventListener?.('message', receive);
    delete owner[DETACH];
    delete owner[ATTACHED];
  };
  worker.addEventListener('message', receive);
  owner[DETACH] = detach;
  owner[ATTACHED] = true;
  return detach;
}
