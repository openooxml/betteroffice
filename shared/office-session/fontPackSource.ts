import { createFontPackResponder, type OptionalFontPackSource } from './optionalFontPacks';

const importOptionalFontPack: OptionalFontPackSource = async (pack) =>
  (await import('@betteroffice/fonts')).optionalFontPackUrls(pack);

export function attachFontPackSource(
  worker: Pick<Worker, 'postMessage' | 'addEventListener'>,
  source = importOptionalFontPack,
): void {
  const responder = createFontPackResponder((message) => worker.postMessage(message), source);
  worker.addEventListener?.('message', (event) => {
    if (responder.receive(event.data)) event.stopImmediatePropagation();
  });
}
