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
  /**
   * The characters to cover. Only the slices whose `unicode-range` includes one of them are
   * fetched, so a later call with newly typed characters fetches only the slices those add.
   * Without it every slice of the face is fetched, and a face in more than 12 slices is skipped.
   */
  text?: string;
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

const parsedRanges = new Map<string, [number, number][]>();

function unicodeRanges(range: string): [number, number][] {
  let ranges = parsedRanges.get(range);
  if (ranges) return ranges;
  ranges = [];
  for (const part of range.split(',')) {
    const match = /^\s*U\+([0-9a-f?]{1,6})(?:-([0-9a-f]{1,6}))?\s*$/i.exec(part);
    if (!match) continue;
    const start = match[1]!;
    ranges.push(match[2] === undefined
      ? [Number.parseInt(start.replace(/\?/g, '0'), 16), Number.parseInt(start.replace(/\?/g, 'f'), 16)]
      : [Number.parseInt(start, 16), Number.parseInt(match[2], 16)]);
  }
  if (parsedRanges.size < MAX_RULES * 4) parsedRanges.set(range, ranges);
  return ranges;
}

/** Whether a CSS `unicode-range` value includes the code point; a rule without one covers every character. */
export function unicodeRangeIncludes(range: string | null, point: number): boolean {
  return range === null || unicodeRanges(range).some(([start, end]) => point >= start && point <= end);
}

/**
 * The provider's slices of the look-alike for an Office `family` in `face`, or in its regular face
 * when the look-alike lacks that weight or style. With `options.text`, only the slices covering one
 * of its characters; without it, every slice, or none for a face in more than 12 slices. Empty when
 * the table has no look-alike or the provider is off. Fetch their bytes with `fetchRemoteFontSlices`.
 */
export async function remoteFontSlices(
  family: string, face: RemoteFaceRequest, options: RemoteFontLoadOptions = {},
): Promise<RemoteFontFaceRule[]> {
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
  if (options.text === undefined) return rules.length > MAX_FACES ? [] : rules;
  const points = new Set<number>();
  for (const character of options.text) points.add(character.codePointAt(0)!);
  return rules.filter((rule) => {
    for (const point of points) if (unicodeRangeIncludes(rule.unicodeRange, point)) return true;
    return false;
  });
}

async function fetchSlice(rule: RemoteFontFaceRule, options: RemoteFontLoadOptions, signal: AbortSignal): Promise<ArrayBuffer> {
  const request = options.fetch ?? globalThis.fetch;
  const response = await request(preferredSource(rule).url, { credentials: 'omit', referrerPolicy: 'no-referrer', signal });
  if (!response.ok) throw new Error(`Remote font failed: ${response.status} ${response.url}`);
  const bytes = await response.arrayBuffer();
  if (bytes.byteLength > MAX_FONT_BYTES) throw new Error(`Remote font too large: ${response.url}`);
  return bytes;
}

/** The font files of `slices`, fetched without credentials or referrer, in the same order. */
export function fetchRemoteFontSlices(
  slices: readonly RemoteFontFaceRule[], options: RemoteFontLoadOptions = {},
): Promise<ArrayBuffer[]> {
  const signal = deadline(options);
  return within(Promise.all(slices.map((rule) => fetchSlice(rule, options, signal))), signal);
}

const remoteFontFaces = new WeakSet<object>();

/** Whether `face` was registered by `loadRemoteFontFaces`, rather than by the host or the browser. */
export function isRemoteFontFace(face: FontFace): boolean {
  return remoteFontFaces.has(face);
}

async function remoteFace(alias: string, rule: RemoteFontFaceRule, bytes: ArrayBuffer): Promise<FontFace> {
  const face = await new FontFace(alias, bytes, {
    weight: rule.weight, style: rule.style, ...(rule.unicodeRange ? { unicodeRange: rule.unicodeRange } : {}),
  }).load();
  remoteFontFaces.add(face);
  return face;
}

const MAX_REGISTERED_SLICES = 1024;
interface RegisteredSlice {
  font: Promise<FontFace>;
  loaded?: FontFace;
  waiting: Set<AbortSignal | null>;
}

const registeredSlices = new WeakMap<FontFaceSet, Map<string, RegisteredSlice>>();

/**
 * Registers the provider's look-alike for an Office family under the document's own family name
 * (`alias`) in `fontSet`. With `options.text`, only the unicode-range slices covering its
 * characters are fetched and registered, and slices registered by an earlier call are reused, so
 * newly typed characters fetch only the slices they add. Without it, every slice of the requested
 * face is fetched and registered once all have loaded, and a look-alike served in more than 12
 * slices (the CJK families) is not loaded. The font files are fetched without credentials or
 * referrer, like the stylesheet. A look-alike without the requested weight or style registers its
 * regular face, which the browser emboldens or slants. Resolves to the faces covering the request,
 * or an empty list when the table has no look-alike, the provider is off or nothing matches.
 * Nothing is registered once `signal` aborts. With `text`, concurrent calls share one fetch per
 * slice: an aborting caller stops waiting while the fetch continues for the others, and a slice
 * every caller abandoned is not registered. The record of registered slices lives with `fontSet`,
 * refetches a slice whose face was removed from the set and holds at most 1024 slices; a call
 * needing more rejects.
 */
export async function loadRemoteFontFaces(
  family: string, alias: string, face: RemoteFaceRequest, fontSet: FontFaceSet,
  options: RemoteFontLoadOptions = {},
): Promise<FontFace[]> {
  const slices = await remoteFontSlices(family, face, options);
  if (slices.length === 0) return [];
  if (options.text === undefined) {
    const bytes = await fetchRemoteFontSlices(slices, options);
    const fonts = await Promise.all(slices.map((rule, index) => remoteFace(alias, rule, bytes[index]!)));
    options.signal?.throwIfAborted();
    for (const font of fonts) fontSet.add(font);
    return fonts;
  }
  let registered = registeredSlices.get(fontSet);
  if (!registered) registeredSlices.set(fontSet, registered = new Map());
  const keyOf = (rule: RemoteFontFaceRule) => JSON.stringify([alias, rule.weight, rule.style, rule.unicodeRange, preferredSource(rule).url]);
  const present = (entry: RegisteredSlice) => !entry.loaded || typeof fontSet.has !== 'function' || fontSet.has(entry.loaded);
  for (const rule of slices) {
    const entry = registered.get(keyOf(rule));
    if (entry && !present(entry)) registered.delete(keyOf(rule));
  }
  const fresh = slices.filter((rule) => !registered.has(keyOf(rule)));
  if (registered.size + fresh.length > MAX_REGISTERED_SLICES) {
    for (const [key, entry] of registered) if (!present(entry)) registered.delete(key);
  }
  if (registered.size + fresh.length > MAX_REGISTERED_SLICES) throw new Error('Remote font slice record is full');
  for (const rule of fresh) {
    const key = keyOf(rule);
    const waiting = new Set<AbortSignal | null>();
    const entry: RegisteredSlice = { waiting, font: fetchSlice(rule, options, deadline(options, true))
      .then((bytes) => remoteFace(alias, rule, bytes))
      .then((font) => {
        const live = [...waiting].some((signal) => signal === null || !signal.aborted);
        waiting.clear();
        if (!live) throw new Error('Remote font slice abandoned by every caller');
        fontSet.add(font);
        return font;
      }) };
    registered.set(key, entry);
    void entry.font.then((font) => { entry.loaded = font; }, () => { if (registered.get(key) === entry) registered.delete(key); });
  }
  const entries = slices.flatMap((rule) => registered.get(keyOf(rule)) ?? []);
  for (const entry of entries) if (!entry.loaded) entry.waiting.add(options.signal ?? null);
  const fonts = Promise.all(entries.map((entry) => entry.font));
  return options.signal ? within(fonts, options.signal) : fonts;
}
