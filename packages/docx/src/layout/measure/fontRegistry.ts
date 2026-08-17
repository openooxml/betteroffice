/** Builds deterministic embedded, bundled, script, and last-resort font chains. */

/** Byte sink of the wasm text engine (`crates/docx-text` FontStore). */
export interface TextEngineFontSink {
  /**
   * Register raw sfnt bytes with the engine; returns its font id.
   * Throws on unparseable bytes.
   */
  registerFont(bytes: Uint8Array): number;
}

/**
 * Script bucket a fallback face provides glyph coverage for. Han text is
 * region-resolved by the caller (SC/TC/JP/KR — see `rustMeasureSource.ts`);
 * kana, Hangul, Arabic and Hebrew detect directly.
 *
 * @public
 */
export type FontScript = 'cjk-sc' | 'cjk-tc' | 'cjk-jp' | 'cjk-kr' | 'arabic' | 'hebrew';

/**
 * Resolver for the bundled metric-compatible set (Carlito↔Calibri,
 * Liberation↔Arial/Times/Courier, …). Returns a lazy byte loader so the
 * (lazily fetched, same-origin) binary is only downloaded when a document
 * actually needs the face, or `undefined` when the family has no bundled
 * substitute.
 *
 * @public
 */
export interface BundledFontProvider {
  /** Resolve a Word family to bundled metric-compatible face byte loaders, or undefined. */
  resolve(family: string, bold: boolean, italic: boolean): (() => Promise<ArrayBuffer>) | undefined;
  /**
   * Optional per-script coverage fallback (Noto CJK/RTL faces). Same loader
   * contract as {@link BundledFontProvider.resolve}; providers without
   * script faces simply omit the method. These faces are coverage fallbacks
   * first, metric approximations second — CJK metric compatibility is much
   * weaker than the Carlito/Calibri class of substitutes.
   */
  resolveScriptFallback?(
    script: FontScript,
    bold: boolean,
    italic: boolean
  ): (() => Promise<ArrayBuffer>) | undefined;
  /**
   * The always-available last-resort base face (broad-coverage Latin —
   * Liberation Sans/Serif per serif-ness). Unlike {@link
   * BundledFontProvider.resolve}, a conforming provider (e.g.
   * `@betteroffice/fonts`'s `resolveLastResortFace`) returns a loader for
   * EVERY family, so the chain this registry builds is guaranteed non-empty
   * and a run never routes to the browser measurer for want of font bytes. The
   * face's metrics are the base font's, not the requested family's — the
   * accepted divergence for a truly-unknown font, in exchange for staying on
   * the native measurement path.
   *
   * Optional so mock/partial providers can omit it; when absent (or returning
   * undefined) an unmapped family still yields an empty chain and the caller
   * browser-falls-back that run — the pre-policy behavior. Same lazy loader
   * contract as {@link BundledFontProvider.resolve}.
   */
  resolveLastResort?(
    family: string,
    bold: boolean,
    italic: boolean
  ): (() => Promise<ArrayBuffer>) | undefined;
}

/**
 * One embedded face extracted from the open document. Structurally identical
 * to `EmbeddedFontFace` (`utils/embeddedFonts.ts`), so the output of
 * `getEmbeddedFontFaces` feeds straight in; duplicated here so this module
 * stays decoupled from the loader.
 *
 * @public
 */
export interface EmbeddedFaceInput {
  /** Word font name the face is registered under (attacker-controlled). */
  family: string;
  /** `'bold'` for the embedBold/embedBoldItalic slots, else `'normal'`. */
  weight: 'normal' | 'bold';
  /** `'italic'` for the embedItalic/embedBoldItalic slots, else `'normal'`. */
  style: 'normal' | 'italic';
  /** De-obfuscated OpenType/TrueType bytes. */
  data: ArrayBuffer;
  /** Whether the source face was subsetted (`w:subsetted`). */
  subsetted: boolean;
}

/**
 * Family keys are matched case-insensitively and whitespace-trimmed —
 * the same normalization `resolveFontFamily` (`utils/fontResolver.ts`)
 * applies to Word font names.
 */
function familyKey(family: string): string {
  return family.trim().toLowerCase();
}

function chainKey(family: string, bold: boolean, italic: boolean): string {
  return `${familyKey(family)}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
}

/** A provider or async factory; `undefined` means no bundled bytes. @public */
export type BundledFontProviderSource =
  | BundledFontProvider
  | (() => Promise<BundledFontProvider | undefined>);

interface ChainResolution {
  ids: number[];
  retryable: boolean;
}

interface ScriptResolution {
  id: number | null;
  retryable: boolean;
}

interface PromiseMemo<K, V> {
  get(key: K): Promise<V> | undefined;
  delete(key: K): boolean;
}

function evictFailedPromise<K, V>(
  memo: PromiseMemo<K, V>,
  key: K,
  pending: Promise<V>,
  failed: (value: V) => boolean
): void {
  pending.then(
    (value) => {
      if (failed(value) && memo.get(key) === pending) memo.delete(key);
    },
    () => {
      if (memo.get(key) === pending) memo.delete(key);
    }
  );
}

export class TextMeasureFontRegistry {
  private readonly sink: TextEngineFontSink;
  private readonly bundledSource: BundledFontProviderSource | undefined;
  private bundledPromise: Promise<BundledFontProvider | undefined> | undefined;
  /** Per-registry so misconfigured editors warn independently. */
  private warnedFallback = false;

  /** Normalized family → embedded faces of the current document. */
  private facesByFamily = new Map<string, EmbeddedFaceInput[]>();
  /**
   * Per-face registration memo, keyed by face object identity so re-feeding
   * the same faces (or sharing one face across several chains) never
   * re-registers its bytes. Failed registrations are evicted. Weak so faces of
   * a closed document can be collected.
   */
  private faceIds = new WeakMap<EmbeddedFaceInput, Promise<number | null>>();
  private bundledIds = new Map<string, Promise<number | null>>();
  /**
   * Last-resort base-face registration memo, keyed by chain key. Kept separate
   * from `bundledIds` so a family's metric-compat face and its (possibly
   * different) last-resort base face don't clobber each other's registration.
   * Document-independent like `bundledIds` — survives `setEmbeddedFaces`, reset
   * by `clear()`. Failed registrations are evicted.
   */
  private lastResortIds = new Map<string, Promise<number | null>>();
  /**
   * Byte-identity dedupe across every registration path. Providers cache
   * their fetches, so the same face requested under several chain keys (or
   * as both a family face and a script fallback) resolves to one buffer —
   * and must produce one engine id, not one copy of the bytes per key.
   * Rejections are evicted. Weak so buffers of a cleared registry can be
   * collected.
   */
  private bufferIds = new WeakMap<ArrayBuffer, Promise<number>>();
  private scriptIds = new Map<FontScript, Promise<ScriptResolution>>();
  /** Settled per-script results for the synchronous view. */
  private scriptResults = new Map<FontScript, ScriptResolution>();
  /** Chain memo — concurrent `getFontIdChain` calls share one resolution. */
  private chains = new Map<string, Promise<ChainResolution>>();
  /** Settled chains for the synchronous view. */
  private chainResults = new Map<string, ChainResolution>();
  /** Bumped on invalidation so stale in-flight resolutions can't repopulate caches. */
  private generation = 0;

  constructor(sink: TextEngineFontSink, opts?: { bundled?: BundledFontProviderSource }) {
    this.sink = sink;
    this.bundledSource = opts?.bundled;
  }

  /** Shares in-flight resolution but evicts misses for retry. */
  private bundled(): Promise<BundledFontProvider | undefined> {
    if (this.bundledPromise === undefined) {
      const source = this.bundledSource;
      const promise = Promise.resolve()
        .then(() => (typeof source === 'function' ? source() : source))
        .catch(() => undefined);
      promise.then((provider) => {
        if (provider === undefined && this.bundledPromise === promise) {
          this.bundledPromise = undefined;
        }
      });
      this.bundledPromise = promise;
    }
    return this.bundledPromise;
  }

  /**
   * Feed the embedded faces extracted from the open document, replacing any
   * previous set and invalidating every cached chain. Font ids already issued
   * by the sink stay valid (the engine keeps the bytes); faces passed again
   * by object identity keep their registration. Script-fallback
   * registrations are document-independent and survive.
   */
  setEmbeddedFaces(faces: EmbeddedFaceInput[]): void {
    this.generation++;
    this.facesByFamily = new Map();
    for (const face of faces) {
      const key = familyKey(face.family);
      const list = this.facesByFamily.get(key);
      if (list) list.push(face);
      else this.facesByFamily.set(key, [face]);
    }
    this.chains = new Map();
    this.chainResults = new Map();
  }

  /**
   * Lazily resolves and registers a font chain. Concurrent calls share work;
   * failed registrations remain retryable.
   */
  getFontIdChain(family: string, bold: boolean, italic: boolean): Promise<number[]> {
    const key = chainKey(family, bold, italic);
    let chain = this.chains.get(key);
    if (!chain) {
      const chains = this.chains;
      const results = this.chainResults;
      const generation = this.generation;
      const pending = this.resolveChain(family, bold, italic);
      pending.then(
        (resolution) => {
          if (generation !== this.generation) return;
          results.set(key, resolution);
          if (resolution.retryable && chains.get(key) === pending) chains.delete(key);
        },
        () => {
          if (chains.get(key) === pending) chains.delete(key);
        }
      );
      chains.set(key, pending);
      chain = pending;
    }
    return chain.then((resolution) => resolution.ids);
  }

  /** Synchronous settled view; retryable results are evicted after one read. */
  getCachedFontIdChain(
    family: string,
    bold: boolean,
    italic: boolean
  ): readonly number[] | undefined {
    const key = chainKey(family, bold, italic);
    const resolution = this.chainResults.get(key);
    if (resolution?.retryable) {
      const results = this.chainResults;
      queueMicrotask(() => {
        if (results.get(key) === resolution) results.delete(key);
      });
    }
    return resolution?.ids;
  }

  /** Resolves unique script fallbacks in order; stable misses cache and failures retry. */
  getScriptFallbackIds(scripts: FontScript[]): Promise<number[]> {
    const unique = [...new Set(scripts)];
    return Promise.all(unique.map((script) => this.resolveScript(script))).then((resolutions) => {
      const out: number[] = [];
      for (const { id } of resolutions) {
        if (id !== null && !out.includes(id)) out.push(id);
      }
      return out;
    });
  }

  /** Synchronous settled view; retryable results are evicted after one pass. */
  getCachedScriptFallbackIds(scripts: FontScript[]): readonly number[] | undefined {
    const results = this.scriptResults;
    const resolutions: Array<[FontScript, ScriptResolution]> = [];
    for (const script of scripts) {
      const resolution = results.get(script);
      if (!resolution) return undefined;
      resolutions.push([script, resolution]);
    }
    const out: number[] = [];
    for (const [script, resolution] of resolutions) {
      const { id } = resolution;
      if (id !== null && !out.includes(id)) out.push(id);
      if (resolution.retryable) {
        queueMicrotask(() => {
          if (results.get(script) === resolution) results.delete(script);
        });
      }
    }
    return out;
  }

  /**
   * Forget every cached chain, face registration, bundled registration and
   * script-fallback registration. Call when the engine/FontStore is
   * recreated — previously issued font ids are invalid then. The embedded
   * face set is retained; pass a new set via `setEmbeddedFaces` when the
   * document changes.
   */
  clear(): void {
    this.generation++;
    this.bundledPromise = undefined;
    this.faceIds = new WeakMap();
    this.bundledIds = new Map();
    this.lastResortIds = new Map();
    this.bufferIds = new WeakMap();
    this.scriptIds = new Map();
    this.scriptResults = new Map();
    this.chains = new Map();
    this.chainResults = new Map();
  }

  private async resolveChain(
    family: string,
    bold: boolean,
    italic: boolean
  ): Promise<ChainResolution> {
    const key = chainKey(family, bold, italic);
    const ids: number[] = [];
    let retryable = false;

    const face = this.pickEmbeddedFace(family, bold, italic);
    if (face) {
      const id = await this.registerFace(face);
      if (id !== null) ids.push(id);
      else retryable = true;
    }

    const bundled = await this.bundled();
    if (bundled === undefined && typeof this.bundledSource === 'function') retryable = true;
    let loader: (() => Promise<ArrayBuffer>) | undefined;
    try {
      loader = bundled?.resolve(family, bold, italic);
    } catch (error) {
      retryable = true;
      console.warn(
        `[fontRegistry] bundled face resolver for "${family}" failed; falling back: ` +
          `${error instanceof Error ? error.message : String(error)}`
      );
    }
    if (loader) {
      const id = await this.registerBundled(key, loader, family);
      if (id !== null && !ids.includes(id)) ids.push(id);
      if (id === null) retryable = true;
    }

    // Terminal link: the always-available last-resort base face. Appended after
    // the embedded + metric-compatible faces so the chain NEVER ends empty — a
    // run whose family has no embedded/bundled match still has real (Liberation)
    // bytes to measure with, keeping the block on the native path instead of the
    // browser fallback. Deduped by engine id (a family whose metric-compat IS
    // the base face, e.g. Arial→Liberation Sans, contributes one id).
    let lastResort: (() => Promise<ArrayBuffer>) | undefined;
    try {
      lastResort = bundled?.resolveLastResort?.(family, bold, italic);
    } catch (error) {
      retryable = true;
      console.warn(
        `[fontRegistry] last-resort resolver for "${family}" failed; falling back: ` +
          `${error instanceof Error ? error.message : String(error)}`
      );
    }
    if (lastResort) {
      const id = await this.registerLastResort(key, lastResort, family);
      if (id !== null && !ids.includes(id)) ids.push(id);
      if (id === null) retryable = true;
    }

    if (ids.length === 0) {
      this.warnFallback(
        family,
        bundled !== undefined,
        loader !== undefined || lastResort !== undefined
      );
    }

    return { ids, retryable };
  }

  /**
   * An empty chain forces browser measurement. Distinguish absent, incomplete,
   * and broken providers so the one warning gives applicable remediation.
   */
  private warnFallback(family: string, providerResolved: boolean, loaderResolved: boolean): void {
    if (this.warnedFallback) return;
    this.warnedFallback = true;
    const consequence =
      'Native measurement is unsupported; browser hosts fall back to measureText and may use ' +
      'different OS fonts, so pagination can vary. Reported once per registry.';
    console.warn(
      loaderResolved
        ? `[fontRegistry] no font bytes for "${family}" — the font provider resolved but every ` +
            `matching face failed to load. ${consequence} If a bundler inlined ` +
            '@betteroffice/fonts, mark it external so its asset URLs resolve.'
        : providerResolved
          ? `[fontRegistry] no font bytes for "${family}" — the configured font provider has ` +
            `no matching face. Add coverage for this family. ${consequence}`
          : `[fontRegistry] no font bytes for "${family}" — no font provider is configured. ` +
            'Install @betteroffice/fonts and pass it with configureDefaultFonts({ fonts }). ' +
            consequence
    );
  }

  /**
   * Exact (weight, style) match first; otherwise the family's regular face —
   * the engine synthesizes bold/italic from regular outlines when asked for a
   * style the document did not embed.
   */
  private pickEmbeddedFace(
    family: string,
    bold: boolean,
    italic: boolean
  ): EmbeddedFaceInput | undefined {
    const faces = this.facesByFamily.get(familyKey(family));
    if (!faces) return undefined;
    const weight = bold ? 'bold' : 'normal';
    const style = italic ? 'italic' : 'normal';
    return (
      faces.find((f) => f.weight === weight && f.style === style) ??
      faces.find((f) => f.weight === 'normal' && f.style === 'normal')
    );
  }

  /** Register raw bytes exactly once per buffer identity (see `bufferIds`). */
  private registerBuffer(bytes: ArrayBuffer): Promise<number> {
    let pending = this.bufferIds.get(bytes);
    if (!pending) {
      const ids = this.bufferIds;
      pending = Promise.resolve().then(() => this.sink.registerFont(new Uint8Array(bytes)));
      evictFailedPromise(ids, bytes, pending, () => false);
      ids.set(bytes, pending);
    }
    return pending;
  }

  private registerFace(face: EmbeddedFaceInput): Promise<number | null> {
    let pending = this.faceIds.get(face);
    if (!pending) {
      const ids = this.faceIds;
      pending = (async () => {
        try {
          return await this.registerBuffer(face.data);
        } catch {
          // A corrupt embedded face (attacker-controlled bytes the engine
          // rejected) must not take down the chain — drop it and let the
          // bundled/browser fallbacks cover the run. console.warn matches
          // the sibling loader's convention (see loadEmbeddedFonts).
          console.warn(
            `[fontRegistry] embedded face "${face.family}" (${face.weight} ${face.style}) ` +
              'was rejected by the text engine; falling back'
          );
          return null;
        }
      })();
      evictFailedPromise(ids, face, pending, (id) => id === null);
      ids.set(face, pending);
    }
    return pending;
  }

  private registerBundled(
    key: string,
    loader: () => Promise<ArrayBuffer>,
    family: string
  ): Promise<number | null> {
    let pending = this.bundledIds.get(key);
    if (!pending) {
      const ids = this.bundledIds;
      pending = (async () => {
        try {
          const bytes = await loader();
          return await this.registerBuffer(bytes);
        } catch (error) {
          console.warn(
            `[fontRegistry] bundled face for "${family}" failed to load or register; ` +
              `falling back: ${error instanceof Error ? error.message : String(error)}`
          );
          return null;
        }
      })();
      evictFailedPromise(ids, key, pending, (id) => id === null);
      ids.set(key, pending);
    }
    return pending;
  }

  /**
   * Register the last-resort base face. Mirrors {@link registerBundled} but
   * uses its own memo so it never collides with metric-compatible registration.
   */
  private registerLastResort(
    key: string,
    loader: () => Promise<ArrayBuffer>,
    family: string
  ): Promise<number | null> {
    let pending = this.lastResortIds.get(key);
    if (!pending) {
      const ids = this.lastResortIds;
      pending = (async () => {
        try {
          const bytes = await loader();
          return await this.registerBuffer(bytes);
        } catch {
          console.warn(
            `[fontRegistry] last-resort base face for "${family}" failed to load or register; ` +
              'falling back'
          );
          return null;
        }
      })();
      evictFailedPromise(ids, key, pending, (id) => id === null);
      ids.set(key, pending);
    }
    return pending;
  }

  private resolveScript(script: FontScript): Promise<ScriptResolution> {
    let pending = this.scriptIds.get(script);
    if (!pending) {
      const ids = this.scriptIds;
      const results = this.scriptResults;
      pending = (async () => {
        let id: number | null = null;
        let retryable = false;
        const bundled = await this.bundled();
        if (bundled === undefined && typeof this.bundledSource === 'function') retryable = true;
        let loader: (() => Promise<ArrayBuffer>) | undefined;
        try {
          loader = bundled?.resolveScriptFallback?.(script, false, false);
        } catch (error) {
          retryable = true;
          console.warn(
            `[fontRegistry] script-fallback resolver for "${script}" failed; falling back: ` +
              `${error instanceof Error ? error.message : String(error)}`
          );
        }
        if (loader) {
          try {
            const bytes = await loader();
            id = await this.registerBuffer(bytes);
          } catch (error) {
            retryable = true;
            console.warn(
              `[fontRegistry] script-fallback face for "${script}" failed to load or register; ` +
                `falling back: ${error instanceof Error ? error.message : String(error)}`
            );
          }
        }
        const resolution = { id, retryable };
        results.set(script, resolution);
        return resolution;
      })();
      evictFailedPromise(ids, script, pending, (resolution) => resolution.retryable);
      ids.set(script, pending);
    }
    return pending;
  }
}
