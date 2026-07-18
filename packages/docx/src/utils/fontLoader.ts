/**
 * Google Fonts Loader
 *
 * Dynamically loads fonts from Google Fonts API with:
 * - Loading state tracking
 * - Duplicate prevention
 * - Callback notifications
 * - Font availability detection
 */

import { resolveFontFamily } from './fontResolver';

// Track loaded fonts to avoid duplicate requests
const loadedFonts = new Set<string>();

// Track fonts currently being loaded
const loadingFonts = new Map<string, Promise<boolean>>();

// Track loaded faces (family|weight) for URL/buffer paths where one
// family can have multiple weights registered independently.
const loadedFaces = new Set<string>();
const loadingFaces = new Map<string, Promise<boolean>>();

// Callbacks to notify when fonts are loaded
const loadCallbacks = new Set<(fonts: string[]) => void>();

// Callbacks to notify when a font fails to load. Adapters subscribe and
// forward to their `onError` prop so library consumers can route into their
// own error tracker (Sentry, Datadog, etc.) instead of filtering the console.
const errorCallbacks = new Set<(error: Error) => void>();

// Track overall loading state
let isLoadingAny = false;

// When false, the automatic Google Fonts fetch path is disabled. Embedders
// that must not egress (e.g. a CSP-locked review surface rendering privileged
// content from embedded font blobs, or any offline host) call
// setGoogleFontsEnabled(false) to suppress the redundant remote fetches at the
// source. Embedded faces (loadFontFromBuffer) and consumer-hosted faces
// (loadFontFromUrl / the `fonts` prop) are unaffected — only the implicit
// Google Fonts lookup in loadFont / loadFontWithMapping is gated.
let googleFontsEnabled = true;

// Families registered through this module's face loaders — raw buffers
// (DOCX-embedded fonts) and consumer URLs (the `fonts` prop). Both are
// routinely subsetted, so "this family renders" does NOT imply full glyph
// coverage. loadFontWithMapping keeps fetching the metric-compatible Google
// equivalent for these families as a glyph-coverage safety net; genuine
// system fonts skip it. Marked synchronously at registration start so an
// in-flight registration already counts as provenance.
const registeredFamilies = new Set<string>();

// Families the canvas probe confirmed as system-satisfied. Kept SEPARATE from
// loadedFonts on purpose: loadedFonts means "a fetch/registration happened"
// and short-circuits loadFont before its options are even looked at, so
// putting probe results there would make a later explicit-weights call a
// silent no-op. Negative probes are NOT cached — a host page's own
// @font-face can make a font appear at any time, and the re-probe is two
// measureText calls on a shared canvas.
const probeSatisfied = new Set<string>();

function reportFontError(error: unknown, context: string): void {
  // Wrap in a fresh Error rather than mutating the original — some Error
  // subclasses (DOMException, frozen objects) have a non-writable .message
  // and assigning to it throws, which would swallow the real load error.
  // Carry the original via `cause` so consumers can still inspect it.
  const origMessage = error instanceof Error ? error.message : String(error);
  const err = new Error(`[font] ${context}: ${origMessage}`, {
    cause: error,
  });

  if (errorCallbacks.size > 0) {
    for (const callback of errorCallbacks) {
      try {
        callback(err);
      } catch (subscriberError) {
        // A bad subscriber must not block the others — but don't silently eat
        // the bug. Surface in dev via console.error so the consumer can fix it.
        console.error('Font error subscriber threw:', subscriberError);
      }
    }
  } else {
    // No subscriber yet — fall back to console so the error is not silently
    // dropped during pre-mount or in non-adapter (headless / SSR) usage.
    console.warn(err.message);
  }
}

function faceKey(
  family: string,
  weight: number | string = 'normal',
  style: 'normal' | 'italic' = 'normal'
): string {
  return `${family.trim()}|${weight}|${style}`;
}

// "A genuine system font satisfies this family" — the decision behind every
// fetch skip. Self-excludes families registered through our face loaders
// (subsetted/partial faces — see registeredFamilies); system fonts that
// render are assumed weight-complete, since OS-bundled families ship full
// sets. The first positive probe also fires onFontsLoaded (microtask,
// matching the fetch path's async timing) — consumers gate ready-state UI
// on that callback, and before the skip existed the fetch guaranteed it.
function satisfiedBySystemFont(family: string): boolean {
  if (registeredFamilies.has(family)) {
    return false;
  }
  if (probeSatisfied.has(family)) {
    return true;
  }
  if (canRenderFont(family)) {
    probeSatisfied.add(family);
    queueMicrotask(() => notifyCallbacks([family]));
    return true;
  }
  return false;
}

// In-flight buffer/URL registrations for a family (loadingFaces is keyed
// `family|weight`).
function inFlightFacePromises(family: string): Promise<boolean>[] {
  const prefix = `${family}|`;
  const pending: Promise<boolean>[] = [];
  for (const [key, promise] of loadingFaces) {
    if (key.startsWith(prefix)) {
      pending.push(promise);
    }
  }
  return pending;
}

// Shared success bookkeeping for the buffer/URL face loaders. One place, so
// the two loaders cannot drift (the next probeSatisfied-style cache line
// added to one but not the other would skew buffer vs URL faces silently).
function markFaceLoaded(key: string, family: string): void {
  loadedFaces.add(key);
  loadedFonts.add(family);
  notifyCallbacks([family]);
}

/**
 * Generate Google Fonts CSS URL for a font family
 *
 * @param fontFamily - The font family name (e.g., "Roboto", "Open Sans")
 * @param weights - Font weights to load (default: 400, 700)
 * @param styles - Font styles to load (default: normal, italic)
 * @returns Google Fonts CSS URL
 */
function getGoogleFontsUrl(
  fontFamily: string,
  weights: number[] = [400, 700],
  styles: ('normal' | 'italic')[] = ['normal', 'italic']
): string {
  // Encode font family name for URL
  const encodedFamily = encodeURIComponent(fontFamily);

  // Build weight/style combinations
  // Format: ital,wght@0,400;0,700;1,400;1,700
  const combinations: string[] = [];

  for (const style of styles) {
    const italVal = style === 'italic' ? 1 : 0;
    for (const weight of weights) {
      combinations.push(`${italVal},${weight}`);
    }
  }

  // Sort and join
  combinations.sort();
  const spec = combinations.join(';');

  return `https://fonts.googleapis.com/css2?family=${encodedFamily}:ital,wght@${spec}&display=swap`;
}

/**
 * Load a font from Google Fonts
 *
 * @param fontFamily - The font family name to load
 * @param options - Optional configuration
 * @returns Promise resolving to true if font loaded successfully, false otherwise
 */
export async function loadFont(
  fontFamily: string,
  options?: {
    weights?: number[];
    styles?: ('normal' | 'italic')[];
  }
): Promise<boolean> {
  // Skip font loading in non-browser environments (Node.js, SSR)
  if (typeof document === 'undefined') {
    return false;
  }

  // Normalize font family name
  const normalizedFamily = fontFamily.trim();

  // Already loaded?
  if (loadedFonts.has(normalizedFamily)) {
    return true;
  }

  // Currently loading? Return existing promise
  const existingLoad = loadingFonts.get(normalizedFamily);
  if (existingLoad) {
    return existingLoad;
  }

  // Already satisfied by a system font — fetching the Google copy would be a
  // redundant round-trip (and a CSP violation in no-egress embedders). Only
  // taken for the default face set: explicit options.weights/options.styles
  // always fetch, because the local probe can only see that *a* face renders,
  // never that a specific weight exists (CSS font matching falls back to the
  // nearest weight instead of failing). Deliberately NOT recorded in
  // loadedFonts — the "Already loaded?" check above ignores options, so
  // recording it there would turn a later explicit-weights call into a silent
  // no-op. Synchronous on purpose: an await here would let concurrent
  // loadFont calls slip past the loadingFonts dedupe above.
  if (!options && satisfiedBySystemFont(normalizedFamily)) {
    return true;
  }

  // A buffer/URL registration for this family may be mid-flight (its
  // waitForFontAvailable can take up to 3s). Snapshot here, in the same
  // synchronous frame as the dedupe checks above.
  const pendingFaces = inFlightFacePromises(normalizedFamily);

  // Remote disabled, not locally satisfied, and no in-flight registration
  // that could change the answer — statically false. Skip the promise
  // machinery instead of re-paying it on every loadDocumentFonts pass.
  if (!googleFontsEnabled && pendingFaces.length === 0) {
    return false;
  }

  // Create load promise
  const loadPromise = (async (): Promise<boolean> => {
    isLoadingAny = true;

    try {
      // Settle in-flight registrations and re-check, so we neither issue the
      // redundant fetch this path exists to skip nor report a spurious
      // failure under googleFontsEnabled=false. Skipped when nothing was in
      // flight — the synchronous checks above already gave the answer.
      if (pendingFaces.length > 0) {
        await Promise.all(pendingFaces);
        if (loadedFonts.has(normalizedFamily)) {
          return true;
        }
        if (!options && satisfiedBySystemFont(normalizedFamily)) {
          return true;
        }
      }

      // Remote fetch disabled (no-egress embedder). The font is not locally
      // available, so don't inject a Google Fonts <link> — report failure and
      // let the caller's CSS fallback stack render with what is available.
      if (!googleFontsEnabled) {
        return false;
      }

      // Generate Google Fonts URL
      const url = getGoogleFontsUrl(normalizedFamily, options?.weights, options?.styles);

      // Create link element
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = url;

      // Wait for load or error, with a 5s timeout. Clear the timer once
      // settled so no handle dangles past the load (keeps test runners and
      // watch cycles from waiting on dead timers).
      const loaded = await new Promise<boolean>((resolve) => {
        const timer = setTimeout(() => resolve(false), 5000);
        link.onload = () => {
          clearTimeout(timer);
          resolve(true);
        };
        link.onerror = () => {
          clearTimeout(timer);
          resolve(false);
        };

        // Append to head
        document.head.appendChild(link);
      });

      if (loaded) {
        // Wait a bit for the font to be available
        await waitForFontAvailable(normalizedFamily, 3000);

        loadedFonts.add(normalizedFamily);

        // Notify callbacks
        notifyCallbacks([normalizedFamily]);

        return true;
      }

      return false;
    } catch (error) {
      reportFontError(error, `failed to load "${normalizedFamily}"`);
      return false;
    } finally {
      loadingFonts.delete(normalizedFamily);

      // Check if still loading any fonts (Google or face-based)
      if (loadingFonts.size === 0 && loadingFaces.size === 0) {
        isLoadingAny = false;
      }
    }
  })();

  loadingFonts.set(normalizedFamily, loadPromise);
  return loadPromise;
}

/**
 * Load multiple fonts from Google Fonts
 *
 * @param families - Array of font family names to load
 * @param options - Optional configuration
 * @returns Promise resolving when all fonts are loaded (or failed)
 */
export async function loadFonts(
  families: string[],
  options?: {
    weights?: number[];
    styles?: ('normal' | 'italic')[];
  }
): Promise<void> {
  // Filter out already loaded fonts
  const toLoad = families.filter((family) => !loadedFonts.has(family.trim()));

  if (toLoad.length === 0) {
    return;
  }

  // Load all fonts in parallel
  await Promise.all(toLoad.map((family) => loadFont(family, options)));
}

/**
 * Check if a font is loaded
 *
 * @param fontFamily - The font family name to check
 * @returns true if the font is loaded, false otherwise
 */
export function isFontLoaded(fontFamily: string): boolean {
  // Probe-satisfied system fonts count as available: onFontsLoaded announces
  // them, so the read API must agree. (loadFont's internal dedupe reads
  // loadedFonts directly — the split only matters there, where a
  // probe-satisfied family must still honor explicit weight requests.)
  const family = fontFamily.trim();
  return loadedFonts.has(family) || probeSatisfied.has(family);
}

/**
 * Enable or disable the editor's automatic Google Fonts lookup.
 *
 * Defaults to enabled. Set to `false` in no-egress embedders (strict CSP,
 * offline) so `loadFont` / `loadFontWithMapping` never inject a
 * `fonts.googleapis.com` stylesheet `<link>`. This gates ONLY the implicit
 * Google Fonts lookup — font URLs you register yourself (the `fonts` prop /
 * `loadFontFromUrl`) are still fetched by the browser, and embedded blobs
 * (`loadFontFromBuffer`) and system fonts still resolve.
 *
 * The flag is page-global (module-level), not per-editor: with multiple
 * editors on one page the last caller wins. Call it before loading documents;
 * disabling does not cancel fetches already in flight. A document font that
 * is neither local nor registered then resolves `loadFont` to `false`
 * silently (no `onFontError`) and renders via its CSS fallback stack.
 *
 * @see isGoogleFontsEnabled
 * @see loadFontFromBuffer
 * @see loadFontFromUrl
 *
 * @public
 */
export function setGoogleFontsEnabled(enabled: boolean): void {
  googleFontsEnabled = enabled;
}

/**
 * Whether the automatic Google Fonts lookup is currently enabled.
 *
 * @see setGoogleFontsEnabled
 *
 * @public
 */
export function isGoogleFontsEnabled(): boolean {
  return googleFontsEnabled;
}

/**
 * Check if any fonts are currently loading
 *
 * @returns true if any fonts are loading, false otherwise
 */
export function isLoading(): boolean {
  return isLoadingAny;
}

/**
 * Get list of all loaded fonts
 *
 * @returns Array of loaded font family names
 */
export function getLoadedFonts(): string[] {
  return Array.from(new Set([...loadedFonts, ...probeSatisfied]));
}

/**
 * Register a callback to be notified when fonts are loaded
 *
 * @param callback - Function to call when fonts are loaded
 * @returns Cleanup function to remove the callback
 */
export function onFontsLoaded(callback: (fonts: string[]) => void): () => void {
  loadCallbacks.add(callback);

  // Return cleanup function
  return () => {
    loadCallbacks.delete(callback);
  };
}

/**
 * Notify all registered callbacks
 */
function notifyCallbacks(fonts: string[]): void {
  for (const callback of loadCallbacks) {
    try {
      callback(fonts);
    } catch (error) {
      reportFontError(error, 'load callback threw');
    }
  }
}

/**
 * Register a callback to be notified when a font fails to load.
 *
 * Adapters subscribe and forward to their `onError` prop. Returns the unsub.
 *
 * @public
 */
export function onFontError(callback: (error: Error) => void): () => void {
  errorCallbacks.add(callback);
  return () => {
    errorCallbacks.delete(callback);
  };
}

/**
 * Wait for a font to be available using the CSS Font Loading API
 *
 * @param fontFamily - The font family to wait for
 * @param timeout - Maximum time to wait in milliseconds
 * @returns Promise resolving when font is available or timeout
 */
async function waitForFontAvailable(fontFamily: string, timeout: number): Promise<boolean> {
  // Use CSS Font Loading API if available
  if ('fonts' in document) {
    try {
      // Try to wait for the font
      const fontFace = `400 16px "${fontFamily}"`;
      await Promise.race([
        document.fonts.load(fontFace),
        new Promise((resolve) => setTimeout(resolve, timeout)),
      ]);

      return document.fonts.check(fontFace);
    } catch {
      // Fall through to fallback
    }
  }

  // Fallback: just wait a bit
  await new Promise((resolve) => setTimeout(resolve, 100));
  return true;
}

// Probe scratchpad. The probe runs once per family per document load, so a
// fresh canvas per call is pure waste — one shared context suffices, and the
// fallback-font widths ('72px sans-serif' over a fixed string) are session
// constants worth memoizing.
let probeContext: CanvasRenderingContext2D | null | undefined;
const fallbackWidthCache = new Map<string, number>();
const PROBE_TEXT = 'abcdefghijklmnopqrstuvwxyz0123456789';

function getProbeContext(): CanvasRenderingContext2D | null {
  if (probeContext === undefined) {
    probeContext = document.createElement('canvas').getContext('2d');
    if (probeContext) {
      probeContext.textBaseline = 'top';
    }
  }
  return probeContext;
}

/**
 * Check if a font is available on the system using canvas measurement
 *
 * Compares text width with the target font vs a known fallback font (and the
 * opposite fallback, for names that collide with the browser defaults). If
 * the widths differ, the font is available. Reuses one shared canvas; the
 * fallback widths are session constants and memoized.
 *
 * @param fontFamily - The font family name to check
 * @param fallbackFont - Fallback font to compare against
 * @returns true if font is available, false otherwise
 */
export function canRenderFont(fontFamily: string, fallbackFont: string = 'sans-serif'): boolean {
  // Skip if we're not in a browser
  if (typeof document === 'undefined') {
    return false;
  }

  const ctx = getProbeContext();
  if (!ctx) {
    return false;
  }

  const measure = (font: string): number => {
    ctx.font = font;
    return ctx.measureText(PROBE_TEXT).width;
  };

  // If widths differ, the custom font (not the fallback) was used.
  const differsFrom = (fallback: string): boolean => {
    let fallbackWidth = fallbackWidthCache.get(fallback);
    if (fallbackWidth === undefined) {
      fallbackWidth = measure(`72px ${fallback}`);
      fallbackWidthCache.set(fallback, fallbackWidth);
    }
    return measure(`72px "${fontFamily}", ${fallback}`) !== fallbackWidth;
  };

  // Check with the primary fallback, then the opposite one (handles the edge
  // case where the font name matches the browser's default sans-serif/serif).
  return (
    differsFrom(fallbackFont) || differsFrom(fallbackFont === 'sans-serif' ? 'serif' : 'sans-serif')
  );
}

// CSS hex escapes (with the mandatory trailing space) for characters that
// either terminate a CSS string token (LF/CR/FF — bad-string-token per
// css-syntax-3) or could close a serialized <style> block (< >).
const CSS_HEX_ESCAPES: Record<string, string> = {
  '<': '\\3c ',
  '>': '\\3e ',
  '\n': '\\a ',
  '\r': '\\d ',
  '\f': '\\c ',
};
const cssStringEscape = (s: string) =>
  s.replace(/["\\]/g, '\\$&').replace(/[<>\n\r\f]/g, (c) => CSS_HEX_ESCAPES[c]);

/**
 * Load a font from a raw buffer (e.g., embedded in DOCX)
 *
 * Call before loading the document when possible: registration marks the
 * family as ours, which is what keeps the metric-compatible Google fallback
 * in play for subsetted faces during the document's font resolution pass.
 * Late registration still self-heals on the next pass. Under
 * `setGoogleFontsEnabled(false)` that fallback fetch is suppressed and
 * glyphs missing from the registered faces render via the CSS stack.
 *
 * @param fontFamily - The font family name
 * @param buffer - Font file buffer (TTF, OTF, WOFF, WOFF2)
 * @param options - Font options
 * @returns Promise resolving when font is loaded
 */
export async function loadFontFromBuffer(
  fontFamily: string,
  buffer: ArrayBuffer,
  options?: {
    weight?: number | string;
    style?: 'normal' | 'italic';
  }
): Promise<boolean> {
  if (typeof document === 'undefined') return false;

  const normalizedFamily = fontFamily.trim();
  const style = options?.style ?? 'normal';
  const key = faceKey(normalizedFamily, options?.weight, style);

  // Provenance: mark before the async work so an in-flight registration
  // already counts as registered (see loadFontWithMapping). Marking a family
  // that then fails to load only means we keep fetching its Google
  // equivalent — the safe direction.
  registeredFamilies.add(normalizedFamily);

  // Face-keyed dedupe so multiple weights of the same family register
  // independently and a prior URL/Google load of the family does not skip
  // this face.
  if (loadedFaces.has(key)) return true;
  const existing = loadingFaces.get(key);
  if (existing) return existing;

  const loadPromise = (async (): Promise<boolean> => {
    isLoadingAny = true;
    try {
      const blob = new Blob([buffer], { type: 'font/ttf' });
      const url = URL.createObjectURL(blob);

      const styleEl = document.createElement('style');
      styleEl.textContent = `
      @font-face {
        font-family: "${cssStringEscape(normalizedFamily)}";
        src: url(${url}) format('truetype');
        font-weight: ${options?.weight ?? 'normal'};
        font-style: ${style};
        font-display: swap;
      }
    `;
      document.head.appendChild(styleEl);

      await waitForFontAvailable(normalizedFamily, 3000);

      markFaceLoaded(key, normalizedFamily);

      return true;
    } catch (error) {
      reportFontError(error, `failed to load "${normalizedFamily}" from buffer`);
      return false;
    } finally {
      loadingFaces.delete(key);
      if (loadingFonts.size === 0 && loadingFaces.size === 0) {
        isLoadingAny = false;
      }
    }
  })();

  loadingFaces.set(key, loadPromise);
  return loadPromise;
}

function guessFontFormat(src: string): string {
  const url = src.split('?')[0].split('#')[0].toLowerCase();
  if (url.endsWith('.woff2')) return 'woff2';
  if (url.endsWith('.woff')) return 'woff';
  if (url.endsWith('.otf')) return 'opentype';
  return 'truetype';
}

/**
 * Load a font face from a URL (woff2, woff, ttf, otf).
 *
 * Injects an `@font-face` rule pointing at the URL. Multiple weights of the
 * same family can be registered independently. Families registered here are
 * treated as potentially subsetted: the editor still fetches their
 * metric-compatible Google equivalent as a glyph-coverage fallback, unless
 * disabled via `setGoogleFontsEnabled(false)`.
 *
 * @param fontFamily - CSS font-family name to expose
 * @param src - URL to the font file
 * @param options - Optional weight
 * @returns Promise resolving to true if the face became available
 *
 * @public
 */
export async function loadFontFromUrl(
  fontFamily: string,
  src: string,
  options?: {
    weight?: number | string;
  }
): Promise<boolean> {
  if (typeof document === 'undefined') return false;

  // Reject URLs containing HTML-breaking characters. The loader writes src
  // into a <style> element's textContent — safe for client rendering, but
  // a serialized document.head (SSR, devtools snapshot) would terminate the
  // style block early on </style>. < and > are never valid in a URL anyway.
  if (/[<>]/.test(src)) {
    reportFontError(
      new Error(`invalid src URL for "${fontFamily}": contains '<' or '>'`),
      'rejected src'
    );
    return false;
  }

  const normalizedFamily = fontFamily.trim();
  const key = faceKey(normalizedFamily, options?.weight);

  // Provenance: subsetted-face safety net — see registeredFamilies.
  registeredFamilies.add(normalizedFamily);

  if (loadedFaces.has(key)) return true;
  const existing = loadingFaces.get(key);
  if (existing) return existing;

  const loadPromise = (async (): Promise<boolean> => {
    isLoadingAny = true;
    try {
      const style = document.createElement('style');
      style.textContent = `
      @font-face {
        font-family: "${cssStringEscape(normalizedFamily)}";
        src: url(${JSON.stringify(src)}) format('${guessFontFormat(src)}');
        font-weight: ${options?.weight ?? 'normal'};
        font-display: swap;
      }
    `;
      document.head.appendChild(style);

      await waitForFontAvailable(normalizedFamily, 3000);

      markFaceLoaded(key, normalizedFamily);

      return true;
    } catch (error) {
      reportFontError(error, `failed to load "${normalizedFamily}" from ${src}`);
      return false;
    } finally {
      loadingFaces.delete(key);
      if (loadingFonts.size === 0 && loadingFaces.size === 0) {
        isLoadingAny = false;
      }
    }
  })();

  loadingFaces.set(key, loadPromise);
  return loadPromise;
}

/**
 * Declarative description of a single font face to register with the editor.
 *
 * Each entry injects one `@font-face` rule pointing at a URL. Multiple
 * entries can share `family` to register distinct weights as separate faces.
 *
 * For Google Fonts, call `loadFont(family)` directly — the `fonts` prop is
 * for fonts the consumer hosts themselves. For raw bytes already in memory
 * (DOCX-embedded fonts, user uploads), call `loadFontFromBuffer(family, buf)`.
 *
 * @public
 */
export interface FontDefinition {
  /**
   * CSS `font-family` name to expose. Match the family name your documents
   * reference; the browser uses this to look up glyphs when text is rendered.
   */
  family: string;
  /**
   * URL to the font file (woff2, woff, ttf, or otf). The loader injects an
   * `@font-face` rule and lets the browser fetch on demand.
   */
  src: string;
  /**
   * CSS `font-weight` for this face. Defaults to `'normal'` (≈400). Pass a
   * number (`400`, `700`) or a CSS keyword (`'bold'`). Required when one
   * `family` registers multiple weights as separate entries.
   */
  weight?: number | string;
}

/**
 * Register a list of custom font faces. Used by the `fonts` prop on
 * `<DocxEditor>` (React + Vue). Idempotent — safe to call on every render.
 *
 * @public
 */
export async function loadFontDefinitions(
  defs: ReadonlyArray<FontDefinition> | undefined
): Promise<void> {
  if (!defs || defs.length === 0) return;
  await Promise.all(
    defs.map((def) => loadFontFromUrl(def.family, def.src, { weight: def.weight }))
  );
}

/**
 * Mapping from common Office/system fonts to Google Fonts equivalents
 *
 * Google Fonts doesn't have exact matches for many Microsoft fonts,
 * but these are close alternatives that work well for document rendering.
 */
export const FONT_MAPPING: Record<string, string> = {
  // Microsoft Office fonts → Google Fonts equivalents
  Calibri: 'Carlito',
  Cambria: 'Caladea',
  Arial: 'Arimo',
  'Times New Roman': 'Tinos',
  'Courier New': 'Cousine',
  Garamond: 'EB Garamond',
  'Book Antiqua': 'EB Garamond',
  Georgia: 'Tinos',
  Verdana: 'Open Sans',
  Tahoma: 'Open Sans',
  'Trebuchet MS': 'Source Sans Pro',
  'Century Gothic': 'Poppins',
  'Franklin Gothic': 'Libre Franklin',
  Palatino: 'EB Garamond',
  'Palatino Linotype': 'EB Garamond',
  'Lucida Sans': 'Open Sans',
  'Segoe UI': 'Open Sans',
  Impact: 'Anton',
  'Comic Sans MS': 'Comic Neue',
  Consolas: 'Inconsolata',
  'Lucida Console': 'Inconsolata',
  Monaco: 'Fira Code',
};

/**
 * Get the Google Fonts equivalent for a font name
 *
 * @param fontName - The original font name from the document
 * @returns The Google Fonts equivalent, or the original name if no mapping exists
 */
export function getGoogleFontEquivalent(fontName: string): string {
  const trimmed = fontName.trim();
  // FONT_MAPPING is the loader's small override table for a few Latin Office
  // fonts. For everything else — notably CJK, and any-case spellings — defer to
  // the single source of truth in fontResolver, whose lookup is case-insensitive
  // and carries the full font→Noto mapping (so the family the loader fetches
  // matches the one fontResolver puts in the CSS fallback stack). Falls back to
  // the raw name when neither maps it.
  return FONT_MAPPING[trimmed] || resolveFontFamily(trimmed).googleFont || trimmed;
}

/**
 * Load a font, automatically mapping to Google Fonts equivalent if needed.
 * If the font needs mapping, also creates a CSS alias so the original font
 * name works in stylesheets.
 *
 * @param fontFamily - The font family name (may be an Office font)
 * @returns Promise resolving to true if font loaded
 */
export async function loadFontWithMapping(fontFamily: string): Promise<boolean> {
  const trimmed = fontFamily.trim();
  const googleFont = getGoogleFontEquivalent(trimmed);

  // Load the Google Font under its own name (no aliasing).
  // The font resolver provides CSS fallback stacks that list both the
  // original DOCX font and the Google equivalent, so the browser will
  // use whichever is available without @font-face aliasing that would
  // hijack Canvas measurements.
  if (googleFont !== trimmed) {
    // A genuine system-installed original has full glyph coverage, so the
    // metric-approximate Google equivalent would never be consulted by the
    // fallback stack — skip its fetch. Registered (subsetted) families keep
    // it as the glyph-coverage safety net — see registeredFamilies. The
    // isFontLoaded leg covers "equivalent already fetched on an earlier
    // pass"; satisfiedBySystemFont self-excludes registered families.
    if (
      (!registeredFamilies.has(trimmed) && isFontLoaded(trimmed)) ||
      satisfiedBySystemFont(trimmed)
    ) {
      return true;
    }
    const result = await loadFont(googleFont);
    if (result) {
      loadedFonts.add(trimmed);
    }
    return result;
  }

  // No mapping needed, load directly
  return loadFont(googleFont);
}

/**
 * Load multiple fonts with automatic mapping to Google Fonts equivalents
 *
 * @param families - Array of font family names
 * @returns Promise resolving when all fonts are loaded
 */
export async function loadFontsWithMapping(families: string[]): Promise<void> {
  // Remove duplicates
  const uniqueFonts = [...new Set(families.map((f) => f.trim()))];
  // Load each font with mapping (creates aliases for Office → Google font mappings)
  await Promise.all(uniqueFonts.map((family) => loadFontWithMapping(family)));
}

/**
 * Preload a list of common document fonts
 *
 * This preloads fonts commonly used in DOCX documents that have
 * Google Fonts equivalents.
 */
export async function preloadCommonFonts(): Promise<void> {
  const commonFonts = [
    'Carlito', // Calibri equivalent
    'Caladea', // Cambria equivalent
    'Arimo', // Arial equivalent
    'Tinos', // Times New Roman equivalent
    'Cousine', // Courier New equivalent
    'EB Garamond', // Garamond equivalent
  ];

  await loadFonts(commonFonts);
}

/**
 * Extract all font families used in a document
 *
 * Uses loose typing to handle any document-like structure.
 *
 * @param document - The parsed document
 * @returns Set of unique font family names
 */
export function extractFontsFromDocument(document: unknown): Set<string> {
  const fonts = new Set<string>();
  if (typeof document !== 'object' || document === null) return fonts;

  // Walk the parsed document/style/table containers, including nested tables,
  // text boxes and header/footer content. Bound the traversal because every
  // parsed DOCX value is attacker-controlled.
  const stack: unknown[] = [document];
  const seen = new WeakSet<object>();
  let visited = 0;
  const MAX_FONT_SCAN_NODES = 100_000;
  while (stack.length > 0 && visited++ < MAX_FONT_SCAN_NODES) {
    const value = stack.pop();
    if (typeof value !== 'object' || value === null) continue;
    if (seen.has(value)) continue;
    seen.add(value);
    if (Array.isArray(value)) {
      for (const child of value) stack.push(child);
      continue;
    }
    const record = value as Record<string, unknown>;
    const family = record.fontFamily;
    if (typeof family === 'object' && family !== null) {
      const slots = family as Record<string, unknown>;
      for (const slot of ['ascii', 'hAnsi', 'eastAsia', 'cs']) {
        const name = slots[slot];
        if (typeof name === 'string' && name.trim()) fonts.add(name.trim());
      }
    }
    // fontTable.xml names and altName aliases are both legitimate authored
    // references; preload both so a run using either spelling resolves.
    if (
      typeof record.name === 'string' &&
      ('altName' in record || 'panose1' in record || 'embedRegular' in record)
    ) {
      if (record.name.trim()) fonts.add(record.name.trim());
      if (typeof record.altName === 'string' && record.altName.trim()) {
        fonts.add(record.altName.trim());
      }
    }
    for (const child of Object.values(record)) stack.push(child);
  }

  return fonts;
}

/**
 * Extract fonts from a document and load them from Google Fonts
 *
 * @param document - The parsed document
 * @returns Promise resolving when fonts are loaded
 */
export async function loadDocumentFonts(document: unknown): Promise<void> {
  const fonts = extractFontsFromDocument(document);

  if (fonts.size === 0) {
    return;
  }

  // Loading document fonts
  await loadFontsWithMapping(Array.from(fonts));
}
