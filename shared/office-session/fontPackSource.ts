import { createFontPackResponder, type OptionalFontPackSource } from './optionalFontPacks';

const importOptionalFontPack: OptionalFontPackSource = async (pack) =>
  (await import('@betteroffice/fonts')).optionalFontPackUrls(pack);
const attached = new WeakSet<object>();

export function attachFontPackSource(
  worker: Pick<Worker, 'postMessage' | 'addEventListener' | 'removeEventListener'>,
  source = importOptionalFontPack,
): (() => void) | undefined {
  if (typeof worker.addEventListener !== 'function' || attached.has(worker)) return;
  const responder = createFontPackResponder((message) => worker.postMessage(message), source);
  const receive = (event: MessageEvent) => {
    if (responder.receive(event.data)) event.stopImmediatePropagation();
  };
  worker.addEventListener('message', receive);
  attached.add(worker);
  return () => worker.removeEventListener?.('message', receive);
}
