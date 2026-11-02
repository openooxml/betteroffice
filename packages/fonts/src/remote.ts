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

function deadline(options: RemoteFontLoadOptions): AbortSignal {
  const timeout = AbortSignal.timeout(options.timeoutMs ?? DEFAULT_TIMEOUT_MS);
  return options.signal && typeof AbortSignal.any === 'function' ? AbortSignal.any([options.signal, timeout]) : timeout;
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
      const signal = deadline(options);
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
  return pending;
}

function cssSource(source: RemoteFontSource): string {
  return `url(${JSON.stringify(source.url)})${source.format ? ` format(${JSON.stringify(source.format)})` : ''}`;
}

const LATIN = 0x41;

function coversLatin(range: string | null): boolean {
  if (!range) return true;
  return range.split(',').some((part) => {
    const [start, end = start] = part.trim().replace(/^u\+/i, '').split('-');
    if (start === undefined || start === '') return false;
    const low = parseInt(start.replace(/\?/g, '0'), 16);
    const high = parseInt((end ?? start).replace(/^u\+/i, '').replace(/\?/g, 'f'), 16);
    return low <= LATIN && LATIN <= high;
  });
}

/**
 * Registers the provider's look-alike for an Office family under the document's own family name
 * (`alias`) in `fontSet`, and resolves once the face covering basic Latin has loaded. Other
 * unicode-range slices stay registered and load when text needs them; `onLateLoad` fires for each.
 * A look-alike without the requested weight or style registers its regular face, which the
 * browser then emboldens or slants.
 * Resolves to the registered faces, or an empty list when the table has no look-alike or the
 * provider is off.
 */
export async function loadRemoteFontFaces(
  family: string, alias: string, face: RemoteFaceRequest, fontSet: FontFaceSet,
  options: RemoteFontLoadOptions & { onLateLoad?: () => void } = {},
): Promise<FontFace[]> {
  const remote = resolveWordFamilySubstitute(family)?.remote;
  if (!remote || (options.provider ?? DEFAULT_REMOTE_FONT_PROVIDER) === 'off') return [];
  const matching = (rules: RemoteFontFaceRule[], wanted: RemoteFaceRequest) =>
    rules.filter((rule) => rule.style === wanted.style && rule.weight === String(wanted.weight));
  let rules = matching(await fetchRemoteFontRules(remote, [face], options), face);
  if (rules.length === 0 && (face.weight !== 400 || face.style !== 'normal')) {
    const regular = { weight: 400, style: 'normal' } as const;
    rules = matching(await fetchRemoteFontRules(remote, [regular], options), regular);
  }
  options.signal?.throwIfAborted();
  const faces = rules.map((rule) => ({
    rule,
    font: new FontFace(alias, rule.sources.map(cssSource).join(', '), {
      weight: rule.weight, style: rule.style, ...(rule.unicodeRange ? { unicodeRange: rule.unicodeRange } : {}),
    }),
  }));
  const eager = faces.filter(({ rule }) => coversLatin(rule.unicodeRange));
  const signal = deadline(options);
  await within(Promise.all(eager.map(({ font }) => font.load())), signal);
  options.signal?.throwIfAborted();
  for (const { font, rule } of faces) {
    fontSet.add(font);
    if (!eager.some((entry) => entry.rule === rule)) void font.loaded.then(() => options.onLateLoad?.(), () => {});
  }
  return faces.map(({ font }) => font);
}
