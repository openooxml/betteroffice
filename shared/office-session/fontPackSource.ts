import { createFontPackResponder, type OptionalFontPackSource } from './optionalFontPacks';
import type { SessionTransport } from './transport';

export const importOptionalFontPack: OptionalFontPackSource = async (pack) =>
  (await import('@betteroffice/fonts')).optionalFontPackUrls(pack);

export function withFontPackSource(transport: SessionTransport, source = importOptionalFontPack): SessionTransport {
  const responder = createFontPackResponder((message) => transport.post(message), source);
  const listeners = new Set<(message: unknown) => void>();
  let unlisten: (() => void) | undefined;
  const stopListening = () => {
    unlisten?.();
    unlisten = undefined;
  };
  return {
    ...transport,
    listen(listener) {
      const forward = (message: unknown) => listener(message);
      listeners.add(forward);
      unlisten ??= transport.listen((message) => {
        if (responder.receive(message)) return;
        for (const receive of [...listeners]) {
          if (listeners.has(receive)) receive(message);
        }
      });
      return () => {
        if (listeners.delete(forward) && listeners.size === 0) stopListening();
      };
    },
    close() {
      responder.dispose();
      listeners.clear();
      stopListening();
      transport.close();
    },
  };
}
