import { readFontBytes } from './bytes';

/**
 * Reads a pack file the way {@link readFontBytes} does, except that a `data:`
 * URL decodes here: bundlers inline small assets as `data:` URLs, and a strict
 * CSP `connect-src` without `data:` refuses a fetch of one.
 */
export async function readFontAsset(file: string, url: URL, timeoutMs?: number): Promise<ArrayBuffer> {
  if (url.protocol !== 'data:') return readFontBytes(file, url, timeoutMs);
  const comma = url.href.indexOf(',');
  if (comma < 0) throw new Error(`Malformed data: URL for bundled font ${file}`);
  const body = decodeURIComponent(url.href.slice(comma + 1));
  if (!/;base64$/i.test(url.href.slice(0, comma))) return new TextEncoder().encode(body).buffer as ArrayBuffer;
  const binary = atob(body);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes.buffer;
}
