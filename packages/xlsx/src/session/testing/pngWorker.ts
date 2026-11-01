import '../worker';
import type { HostMessage } from '../../../../../shared/office-session/protocol';

const fetchAsset = globalThis.fetch;
globalThis.fetch = ((input, init) => {
  const url = input instanceof Request ? input.url : String(input);
  self.postMessage({ protocol: 1, kind: 'event', name: 'png-test-fetch', payload: url } satisfies HostMessage);
  return fetchAsset(input, init);
}) as typeof fetch;
