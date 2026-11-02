import { resolveWordFamilySubstitute } from './substitutes';

export {
  WORD_FAMILY_SUBSTITUTES, resolveWordFamilySubstitute,
  type GenericFontFamily, type WordFamilySubstitute,
} from './substitutes';

/**
 * Where fonts without a bundled metric-compatible face are fetched from:
 * `'bunny'` (fonts.bunny.net, the default), `'google'` (fonts.googleapis.com),
 * `{ cssUrl }` for any Google-compatible CSS endpoint (for example your own
 * pull zone), or `'off'` to never fetch.
 */
export type RemoteFontProvider = 'bunny' | 'google' | 'off' | { cssUrl: string | URL };

export const DEFAULT_REMOTE_FONT_PROVIDER: RemoteFontProvider = 'bunny';

export interface RemoteFaceRequest {
  weight: number;
  style: 'normal' | 'italic';
}

export interface RemoteFontSource {
  url: string;
  format: string | null;
}

export interface RemoteFontFaceRule {
  family: string;
  weight: string;
  style: string;
  unicodeRange: string | null;
  sources: RemoteFontSource[];
}

export interface RemoteFontLoadOptions {
  provider?: RemoteFontProvider;
  signal?: AbortSignal;
  fetch?: typeof globalThis.fetch;
  /** Gives up on a stylesheet or face that has not arrived after this long. Defaults to 5000 ms. */
  timeoutMs?: number;
}

const PROVIDER_CSS: Record<'bunny' | 'google', string> = {
  bunny: 'https://fonts.bunny.net/css',
  google: 'https://fonts.googleapis.com/css',
};

const FAMILY_NAME = /^[\p{L}\p{N}][\p{L}\p{N} .\-]{0,63}$/u;
const MAX_STYLESHEET_BYTES = 512 * 1024;
const MAX_RULES = 512;
const DEFAULT_TIMEOUT_MS = 5000;

function deadline(options: RemoteFontLoadOptions, shared = false): AbortSignal {
  const timeout = AbortSignal.timeout(options.timeoutMs ?? DEFAULT_TIMEOUT_MS);
  const caller = shared ? undefined : options.signal;
  if (!caller) return timeout;
  if (typeof AbortSignal.any === 'function') return AbortSignal.any([caller, timeout]);
  const controller = new AbortController();
  for (const signal of [caller, timeout]) {
    if (signal.aborted) controller.abort(signal.reason);
    else signal.addEventListener('abort', () => controller.abort(signal.reason), { once: true });
  }
  return controller.signal;
}

function within<T>(work: Promise<T>, signal: AbortSignal): Promise<T> {
  if (signal.aborted) return Promise.reject(signal.reason);
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(signal.reason);
    signal.addEventListener('abort', abort, { once: true });
    work.then(resolve, reject).finally(() => signal.removeEventListener('abort', abort));
  });
}

function providerBase(provider: RemoteFontProvider): URL | null {
  if (provider === 'off') return null;
  const base = new URL(typeof provider === 'object' ? String(provider.cssUrl) : PROVIDER_CSS[provider]);
  if (base.protocol !== 'https:' && !(base.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(base.hostname))) {
    throw new TypeError(`Remote font CSS must be served over https: ${base.href}`);
  }
  base.search = '';
  base.hash = '';
  return base;
}

/** The provider's stylesheet URL for `family` in the requested faces, or null when the provider is off. */
export function remoteFontStylesheetUrl(
  family: string, faces: readonly RemoteFaceRequest[] = [{ weight: 400, style: 'normal' }],
  provider: RemoteFontProvider = DEFAULT_REMOTE_FONT_PROVIDER,
): string | null {
  const base = providerBase(provider);
  if (!base) return null;
  const name = family.trim();
  if (!FAMILY_NAME.test(name)) throw new TypeError(`Invalid remote font family: ${JSON.stringify(family)}`);
  const variants = [...new Set(faces.map((face) => `${face.weight}${face.style === 'italic' ? 'i' : ''}`))].sort();
  if (variants.length === 0) throw new TypeError('At least one remote face is required');
  return `${base.href}?family=${encodeURIComponent(name).replace(/%20/g, '+')}:${variants.join(',')}&display=swap`;
}

function declaration(block: string, property: string): string | null {
  const match = new RegExp(`(?:^|[;{\\s])${property}\\s*:\\s*([^;}]*)`, 'i').exec(block);
  return match ? match[1]!.trim() : null;
}

function unquote(value: string): string {
  const trimmed = value.trim();
  return /^(['"]).*\1$/.test(trimmed) ? trimmed.slice(1, -1) : trimmed;
}

/** Parses the `@font-face` rules of a provider stylesheet; source URLs resolve against `base` and must be https. */
export function parseRemoteFontCss(css: string, base: string | URL): RemoteFontFaceRule[] {
  const rules: RemoteFontFaceRule[] = [];
  const stylesheet = new URL(String(base));
  for (const match of css.replace(/\/\*[\s\S]*?\*\//g, '').matchAll(/@font-face\s*\{([^}]*)\}/gi)) {
    if (rules.length >= MAX_RULES) break;
    const block = match[1]!;
    const family = declaration(block, 'font-family');
    const src = declaration(block, 'src');
    if (!family || !src) continue;
    const sources: RemoteFontSource[] = [];
    for (const part of src.matchAll(/url\(\s*(['"]?)([^'")\s]+)\1\s*\)(?:\s*format\(\s*(['"]?)([^'")]+)\3\s*\))?/gi)) {
      const url = new URL(part[2]!, stylesheet);
      if (url.protocol !== 'https:' && url.origin !== stylesheet.origin) continue;
      sources.push({ url: url.href, format: part[4]?.toLowerCase() ?? null });
    }
    if (sources.length === 0) continue;
    rules.push({
      family: unquote(family), weight: declaration(block, 'font-weight') ?? '400',
      style: declaration(block, 'font-style') ?? 'normal',
      unicodeRange: declaration(block, 'unicode-range'), sources,
    });
  }
  return rules;
}

const stylesheets = new Map<string, Promise<RemoteFontFaceRule[]>>();

/** The `@font-face` rules the provider serves for `family`, cached per stylesheet URL. */
export function fetchRemoteFontRules(
  family: string, faces: readonly RemoteFaceRequest[], options: RemoteFontLoadOptions = {},
): Promise<RemoteFontFaceRule[]> {
  const url = remoteFontStylesheetUrl(family, faces, options.provider ?? DEFAULT_REMOTE_FONT_PROVIDER);
  if (!url) return Promise.resolve([]);
  let pending = stylesheets.get(url);
  if (!pending) {
    const request = options.fetch ?? globalThis.fetch;
    pending = (async () => {
      const signal = deadline(options, true);
      const response = await within(request(url, { credentials: 'omit', referrerPolicy: 'no-referrer', signal }), signal);
      if (!response.ok) throw new Error(`Remote font stylesheet failed: ${response.status} ${url}`);
      const css = await within(response.text(), signal);
      if (css.length > MAX_STYLESHEET_BYTES) throw new Error(`Remote font stylesheet too large: ${url}`);
      return parseRemoteFontCss(css, response.url || url);
    })();
    stylesheets.set(url, pending);
    const settled = pending;
    void settled.catch(() => { if (stylesheets.get(url) === settled) stylesheets.delete(url); });
  }
  return options.signal ? within(pending, options.signal) : pending;
}

const MAX_FONT_BYTES = 8 * 1024 * 1024;
const MAX_FACES = 12;
const PREFERRED_FORMATS = ['woff2', 'woff', 'truetype', 'opentype'];

function preferredSource(rule: RemoteFontFaceRule): RemoteFontSource {
  for (const format of PREFERRED_FORMATS) {
    const source = rule.sources.find((entry) => entry.format === format);
    if (source) return source;
  }
  return rule.sources[0]!;
}

/**
 * Registers the provider's look-alike for an Office family under the document's own family name
 * (`alias`) in `fontSet`, once every unicode-range slice of the requested face has loaded. The
 * font files are fetched without credentials or referrer, like the stylesheet. A look-alike
 * served in more than 12 slices (the CJK families) is not loaded. A look-alike without the
 * requested weight or style registers its regular face, which the browser emboldens or slants.
 * Resolves to the registered faces, or an empty list when the table has no look-alike, the
 * provider is off or the face is too large; nothing is registered once `signal` aborts.
 */
export async function loadRemoteFontFaces(
  family: string, alias: string, face: RemoteFaceRequest, fontSet: FontFaceSet,
  options: RemoteFontLoadOptions = {},
): Promise<FontFace[]> {
  const remote = resolveWordFamilySubstitute(family)?.remote;
  if (!remote || (options.provider ?? DEFAULT_REMOTE_FONT_PROVIDER) === 'off') return [];
  const matching = (rules: RemoteFontFaceRule[], wanted: RemoteFaceRequest) =>
    rules.filter((rule) => rule.style === wanted.style && rule.weight === String(wanted.weight));
  let rules = matching(await fetchRemoteFontRules(remote, [face], options), face);
  options.signal?.throwIfAborted();
  if (rules.length === 0 && (face.weight !== 400 || face.style !== 'normal')) {
    const regular = { weight: 400, style: 'normal' } as const;
    rules = matching(await fetchRemoteFontRules(remote, [regular], options), regular);
  }
  options.signal?.throwIfAborted();
  if (rules.length === 0 || rules.length > MAX_FACES) return [];
  const signal = deadline(options);
  const request = options.fetch ?? globalThis.fetch;
  const fonts = await within(Promise.all(rules.map(async (rule) => {
    const response = await request(preferredSource(rule).url,
      { credentials: 'omit', referrerPolicy: 'no-referrer', signal });
    if (!response.ok) throw new Error(`Remote font failed: ${response.status} ${response.url}`);
    const bytes = await response.arrayBuffer();
    if (bytes.byteLength > MAX_FONT_BYTES) throw new Error(`Remote font too large: ${response.url}`);
    return new FontFace(alias, bytes, {
      weight: rule.weight, style: rule.style, ...(rule.unicodeRange ? { unicodeRange: rule.unicodeRange } : {}),
    }).load();
  })), signal);
  options.signal?.throwIfAborted();
  for (const font of fonts) fontSet.add(font);
  return fonts;
}
