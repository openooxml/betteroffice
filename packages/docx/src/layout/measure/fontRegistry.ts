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

/**
 * Either the provider itself, or a factory resolving one asynchronously (the
 * default path dynamic-imports `@betteroffice/fonts`, so it cannot be
 * synchronous). A factory resolving `undefined` means no bundled bytes exist.
 *
 * @public
 */
export type BundledFontProviderSource =
  | BundledFontProvider
  | (() => Promise<BundledFontProvider | undefined>);

export class TextMeasureFontRegistry {
  private readonly sink: TextEngineFontSink;
  private readonly bundledSource: BundledFontProviderSource | undefined;
  /** In-flight or successful provider resolution; misses are retried by later chains. */
  private bundledPromise: Promise<BundledFontProvider | undefined> | undefined;
  /**
   * Whether this registry has already reported a native-measurement fallback.
   * Scoped per registry rather than per module so two editors with different
   * providers each report their own state, and so one misconfigured editor is
   * not silenced by a healthy one.
   */
  private warnedFallback = false;

  /** Normalized family → embedded faces of the current document. */
  private facesByFamily = new Map<string, EmbeddedFaceInput[]>();
  /**
   * Per-face registration memo, keyed by face object identity so re-feeding
   * the same faces (or sharing one face across several chains) never
   * re-registers its bytes. `null` = registration failed (memoized too — a
   * corrupt face stays corrupt). Weak so faces of a closed document can be
   * collected.
   */
  private faceIds = new WeakMap<EmbeddedFaceInput, Promise<number | null>>();
  /** Bundled registration memo, keyed by chain key. `null` = load/registration failed. */
  private bundledIds = new Map<string, Promise<number | null>>();
  /**
   * Last-resort base-face registration memo, keyed by chain key. Kept separate
   * from `bundledIds` so a family's metric-compat face and its (possibly
   * different) last-resort base face don't clobber each other's registration.
   * Document-independent like `bundledIds` — survives `setEmbeddedFaces`, reset
   * by `clear()`. `null` = no face / load failed.
   */
  private lastResortIds = new Map<string, Promise<number | null>>();
  /**
   * Byte-identity dedupe across every registration path. Providers cache
   * their fetches, so the same face requested under several chain keys (or
   * as both a family face and a script fallback) resolves to one buffer —
   * and must produce one engine id, not one copy of the bytes per key.
   * Rejections stay memoized (corrupt bytes stay corrupt). Weak so buffers
   * of a cleared registry can be collected.
   */
  private bufferIds = new WeakMap<ArrayBuffer, Promise<number>>();
  /** Per-script fallback registration memo. `null` = no face / failed. */
  private scriptIds = new Map<FontScript, Promise<number | null>>();
  /** Settled per-script results for the synchronous view. */
  private scriptResults = new Map<FontScript, number | null>();
  /** Chain memo — concurrent `getFontIdChain` calls share one resolution. */
  private chains = new Map<string, Promise<number[]>>();
  /** Settled chains for the synchronous view. */
  private chainResults = new Map<string, readonly number[]>();
  /** Bumped on invalidation so stale in-flight resolutions can't repopulate caches. */
  private generation = 0;

  constructor(sink: TextEngineFontSink, opts?: { bundled?: BundledFontProviderSource }) {
    this.sink = sink;
    this.bundledSource = opts?.bundled;
  }

  /** Share an in-flight provider resolution and retain only a successful result. */
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
   * Ordered font-id chain for a (family, bold, italic) request — see the
   * module doc for the chain order. Lazily registers bytes with the sink on
   * first request and memoizes both per-face registrations and whole chains,
   * so concurrent callers trigger exactly one `registerFont` per face.
   * Resolves to an empty array when neither an embedded nor a bundled face
   * exists — the caller must browser-fallback that run.
   */
  getFontIdChain(family: string, bold: boolean, italic: boolean): Promise<number[]> {
    const key = chainKey(family, bold, italic);
    let chain = this.chains.get(key);
    if (!chain) {
      chain = this.resolveChain(key, family, bold, italic);
      this.chains.set(key, chain);
    }
    return chain;
  }

  /**
   * Synchronous view of an already-resolved chain (for measurement cache
   * keys); `undefined` while unresolved or never requested.
   */
  getCachedFontIdChain(
    family: string,
    bold: boolean,
    italic: boolean
  ): readonly number[] | undefined {
    return this.chainResults.get(chainKey(family, bold, italic));
  }

  /**
   * Resolve (and lazily register) the script-fallback faces for `scripts`,
   * in input order, deduplicated. A script the provider has no face for —
   * or whose face fails to load/register — contributes nothing; like the
   * per-family chains this is failure-tolerant and memoized per script, so
   * a broken loader is not re-hit on every layout pass.
   */
  getScriptFallbackIds(scripts: FontScript[]): Promise<number[]> {
    const unique = [...new Set(scripts)];
    return Promise.all(unique.map((script) => this.resolveScript(script))).then((ids) => {
      const out: number[] = [];
      for (const id of ids) {
        if (id !== null && !out.includes(id)) out.push(id);
      }
      return out;
    });
  }

  /**
   * Synchronous view of the script-fallback ids for `scripts`; `undefined`
   * while ANY of them is still unresolved (the caller should fall back and
   * let `prepareFonts` warm the set). Settled misses contribute nothing.
   */
  getCachedScriptFallbackIds(scripts: FontScript[]): readonly number[] | undefined {
    const out: number[] = [];
    for (const script of scripts) {
      if (!this.scriptResults.has(script)) return undefined;
      const id = this.scriptResults.get(script)!;
      if (id !== null && !out.includes(id)) out.push(id);
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
    key: string,
    family: string,
    bold: boolean,
    italic: boolean
  ): Promise<number[]> {
    const generation = this.generation;
    const ids: number[] = [];

    const face = this.pickEmbeddedFace(family, bold, italic);
    if (face) {
      const id = await this.registerFace(face);
      if (id !== null) ids.push(id);
    }

    const bundled = await this.bundled();
    const loader = bundled?.resolve(family, bold, italic);
    if (loader) {
      const id = await this.registerBundled(key, loader, family);
      if (id !== null && !ids.includes(id)) ids.push(id);
    }

    // Terminal link: the always-available last-resort base face. Appended after
    // the embedded + metric-compatible faces so the chain NEVER ends empty — a
    // run whose family has no embedded/bundled match still has real (Liberation)
    // bytes to measure with, keeping the block on the native path instead of the
    // browser fallback. Deduped by engine id (a family whose metric-compat IS
    // the base face, e.g. Arial→Liberation Sans, contributes one id).
    const lastResort = bundled?.resolveLastResort?.(family, bold, italic);
    if (lastResort) {
      const id = await this.registerLastResort(key, lastResort, family);
      if (id !== null && !ids.includes(id)) ids.push(id);
    }

    if (ids.length === 0) {
      this.warnFallback(
        family,
        bundled !== undefined,
        loader !== undefined || lastResort !== undefined
      );
    }

    // Only publish into the sync view if nothing invalidated us mid-flight.
    if (generation === this.generation) this.chainResults.set(key, ids);
    return ids;
  }

  /**
   * An empty chain makes the native engine return `UNSUPPORTED`; browser hosts
   * then measure the block with `measureText`, which may consult OS fonts.
   * Report this once for the first affected family rather than once per style.
   *
   * Distinguish a failed loader from a partial provider with no matching face,
   * and both from an absent default package, so each audience gets applicable
   * guidance.
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
          : `[fontRegistry] no font bytes for "${family}" — no font provider resolved. Install ` +
            `@betteroffice/fonts. ${consequence}`
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
      pending = Promise.resolve().then(() => this.sink.registerFont(new Uint8Array(bytes)));
      // Every caller handles the rejection; this guard only keeps a memoized
      // failure from surfacing as an unhandled-rejection warning.
      pending.catch(() => {});
      this.bufferIds.set(bytes, pending);
    }
    return pending;
  }

  private registerFace(face: EmbeddedFaceInput): Promise<number | null> {
    let pending = this.faceIds.get(face);
    if (!pending) {
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
      this.faceIds.set(face, pending);
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
      pending = (async () => {
        try {
          const bytes = await loader();
          return await this.registerBuffer(bytes);
        } catch (error) {
          // Failed fetch or unparseable bundled face: memoized as a miss for
          // this registry's lifetime (clear() resets), so a broken loader is
          // not re-hit on every measurement. The message carries the actionable
          // part — this is the path a CJK family reaches when the add-on is
          // missing, since SimHei/Meiryo and friends map through
          // `metricCompatWith` rather than the script-fallback chain.
          console.warn(
            `[fontRegistry] bundled face for "${family}" failed to load or register; ` +
              `falling back: ${error instanceof Error ? error.message : String(error)}`
          );
          return null;
        }
      })();
      this.bundledIds.set(key, pending);
    }
    return pending;
  }

  /**
   * Register (once per chain key) the last-resort base face. Mirrors
   * {@link registerBundled} but keyed in its own memo so it never collides with
   * the family's metric-compat registration.
   */
  private registerLastResort(
    key: string,
    loader: () => Promise<ArrayBuffer>,
    family: string
  ): Promise<number | null> {
    let pending = this.lastResortIds.get(key);
    if (!pending) {
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
      this.lastResortIds.set(key, pending);
    }
    return pending;
  }

  private resolveScript(script: FontScript): Promise<number | null> {
    let pending = this.scriptIds.get(script);
    if (!pending) {
      // Capture the map so a clear() mid-flight publishes into the orphaned
      // instance instead of resurrecting stale ids in the fresh one.
      const results = this.scriptResults;
      pending = (async () => {
        let id: number | null = null;
        const bundled = await this.bundled();
        const loader = bundled?.resolveScriptFallback?.(script, false, false);
        if (loader) {
          try {
            const bytes = await loader();
            id = await this.registerBuffer(bytes);
          } catch (error) {
            // The message carries the actionable part for the CJK buckets —
            // "install @betteroffice/fonts-cjk" — so surface it.
            console.warn(
              `[fontRegistry] script-fallback face for "${script}" failed to load or register; ` +
                `falling back: ${error instanceof Error ? error.message : String(error)}`
            );
          }
        }
        results.set(script, id);
        return id;
      })();
      this.scriptIds.set(script, pending);
    }
    return pending;
  }
}
