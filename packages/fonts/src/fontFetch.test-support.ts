import { spyOn } from 'bun:test';

export function spyOnFontFetch(implementation: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>) {
  const fetch = Object.assign(
    (input: RequestInfo | URL, init?: RequestInit) => implementation(input, init),
    { preconnect: globalThis.fetch.preconnect },
  );
  return spyOn(globalThis, 'fetch').mockImplementation(fetch);
}
