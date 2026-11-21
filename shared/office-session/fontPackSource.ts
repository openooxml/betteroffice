import { createFontPackResponder, isFontPackRequest, type OptionalFontPackSource } from './optionalFontPacks';
import type { SessionTransport } from './transport';

export const importOptionalFontPack: OptionalFontPackSource = async (pack) =>
  (await import('@betteroffice/fonts')).optionalFontPackUrls(pack);

export function withFontPackSource(transport: SessionTransport, source = importOptionalFontPack): SessionTransport {
  const responder = createFontPackResponder((message) => transport.post(message), source);
  const unlisten = transport.listen((message) => { responder.receive(message); });
  return {
    ...transport,
    listen: (listener) => transport.listen((message) => { if (!isFontPackRequest(message)) listener(message); }),
    close() { responder.dispose(); unlisten(); transport.close(); },
  };
}
