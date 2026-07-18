/**
 * The mandatory Rust pagination source — the seam through which EVERY
 * pagination pass (full, incremental-trigger, footnote pass 1, and each
 * footnote-stabilization iteration in `layout/regions/footnoteLayout.ts`)
 * runs the wasm `docx-layout` engine. There is no TypeScript pagination
 * kernel anymore; this seam is the only way a `Layout` is produced.
 *
 * Contract (pinned by the golden corpus — `__golden__/rustParity.test.ts`
 * proves the engine reproduces every golden byte-for-byte, and
 * `__golden__/rustPaginationSeam.test.ts` proves this seam adds zero drift
 * on top):
 *  - Input marshaling is exactly the golden convention:
 *    `JSON.stringify({ measured, options }, mapReplacer)` where the replacer
 *    turns `Map`s (`options.footnoteReservedHeights`) into plain objects
 *    keyed by decimal page number — the shape the Rust `LayoutOptions`
 *    deserializer expects.
 *  - `layout_document_json` in, `Layout` JSON out. Engine failures THROW —
 *    the layout pipeline surfaces the error (`onError` / console) and keeps
 *    the previous committed layout; nothing falls back to a hidden kernel.
 *  - The wasm engine is synchronous once loaded. Adapters gate layout passes
 *    on {@link RustPaginationSource.isReady} / `whenReady()` (the same
 *    defer-and-rerun pattern as the Rust measurement readiness gate), so
 *    `paginate` never runs against an unloaded engine. The wasm module is
 *    shared with the measurement engine, so in practice it is already
 *    resolved by the time the first measured pass reaches pagination.
 *
 * @packageDocumentation
 * @public
 */

import type { Layout, LayoutOptions } from './types';
import type { MeasuredBlock } from './measuredBlock';

/**
 * The narrow seam `editor/computeLayout.ts` consumes: a synchronous
 * pagination call over the measured blocks + options, plus the readiness
 * surface adapters use to defer layout passes until the wasm engine has
 * loaded. `paginate` throws if called before readiness or when the engine
 * rejects the input — there is no fallback engine.
 *
 * @public
 */
export interface LayoutPaginationSource {
  /** Paginate through the loaded Rust engine. Throws when not ready or on engine failure. */
  paginate(measured: MeasuredBlock[], options: LayoutOptions): Layout;
  /** Whether the wasm engine has finished its lazy load. */
  isReady(): boolean;
  /** Resolves when the engine is loaded; rejects if the load failed. */
  whenReady(): Promise<void>;
}

/**
 * Thin sync facade over the wasm pagination export (`layout_document_json`).
 * Injectable so unit tests can script responses without loading wasm — the
 * same pattern as the measurement seam's `RustTextEngine`.
 *
 * @public
 */
export interface RustLayoutEngine {
  /** `{ measured, options }` JSON in, `Layout` JSON out. Throws on any input the engine rejects. */
  layoutDocumentJson(input: string): string;
}

let enginePromise: Promise<RustLayoutEngine> | null = null;

/**
 * Lazily load the embedded wasm module and expose its pagination surface as
 * a {@link RustLayoutEngine}. The dynamic `import()` keeps the ~800KB base64
 * module out of the synchronous require graph; the module — and therefore
 * the wasm instantiation — is shared with the measurement seam
 * (`getRustTextEngine`), so only one load happens per session.
 *
 * @public
 */
export function getRustLayoutEngine(): Promise<RustLayoutEngine> {
  enginePromise ??= import('../wasm/index').then(async (m) => {
    await m.preloadLayoutWasm();
    return {
    layoutDocumentJson: m.layoutDocumentJson,
  };
  });
  return enginePromise;
}

/** Counters for observability (passes paginated / engine errors). @public */
export interface RustPaginationStats {
  /** Layout passes paginated by the Rust engine. */
  rustPaginated: number;
  /** Passes where the engine threw (the pass produced no layout). */
  engineErrors: number;
}

/** Options for {@link createRustPaginationSource}. @public */
export interface RustPaginationSourceOptions {
  /** Pre-loaded engine (tests). When set, the source is ready immediately. */
  engine?: RustLayoutEngine;
  /** Engine loader override (tests). Defaults to {@link getRustLayoutEngine}. */
  loadEngine?: () => Promise<RustLayoutEngine>;
}

/**
 * The mandatory pagination source. Create with
 * {@link createRustPaginationSource}; see the module doc for the contract.
 *
 * @public
 */
export interface RustPaginationSource extends LayoutPaginationSource {
  /** Snapshot of the pagination counters (never reset implicitly). */
  getStats(): RustPaginationStats;
}

/**
 * Maps don't survive JSON.stringify (they become `{}`); options carry
 * `footnoteReservedHeights: Map<number, number>`, marshaled as a plain object
 * keyed by decimal page number — the golden-fixture convention the Rust
 * `LayoutOptions` deserializer expects (`__golden__/rustParity.test.ts`,
 * `scripts/export-golden-fixtures.ts`).
 */
function mapReplacer(_key: string, value: unknown): unknown {
  return value instanceof Map ? Object.fromEntries(value) : value;
}

/** Parse + shape-check the engine's JSON so a malformed payload fails loudly. */
function parseLayout(json: string): Layout {
  const parsed = JSON.parse(json) as Layout;
  if (
    parsed === null ||
    typeof parsed !== 'object' ||
    !Array.isArray(parsed.pages) ||
    typeof parsed.pageSize?.w !== 'number' ||
    typeof parsed.pageSize?.h !== 'number'
  ) {
    throw new Error('rust pagination returned a malformed Layout');
  }
  return parsed;
}

/**
 * Build a {@link RustPaginationSource}. The wasm load starts immediately at
 * construction; adapters defer layout passes until `isReady()` and re-run
 * the pipeline off `whenReady()` — mirroring the Rust measurement readiness
 * gate, and sharing the same underlying wasm module load.
 *
 * @public
 */
export function createRustPaginationSource(
  opts: RustPaginationSourceOptions = {}
): RustPaginationSource {
  let engine: RustLayoutEngine | null = opts.engine ?? null;
  const stats: RustPaginationStats = { rustPaginated: 0, engineErrors: 0 };

  const ready: Promise<void> = engine
    ? Promise.resolve()
    : (opts.loadEngine ?? getRustLayoutEngine)().then((loaded) => {
        engine = loaded;
      });
  // The adapter observes failures through whenReady(); avoid an unhandled
  // rejection when nothing has subscribed yet.
  ready.catch(() => {});

  return {
    isReady(): boolean {
      return engine !== null;
    },

    whenReady(): Promise<void> {
      return ready;
    },

    getStats(): RustPaginationStats {
      return { ...stats };
    },

    paginate(measured: MeasuredBlock[], options: LayoutOptions): Layout {
      if (!engine) {
        throw new Error(
          '[rustPagination] paginate called before the wasm engine loaded — ' +
            'layout passes must be gated on isReady()/whenReady()'
        );
      }
      let layout: Layout;
      try {
        const input = JSON.stringify({ measured, options }, mapReplacer);
        layout = parseLayout(engine.layoutDocumentJson(input));
      } catch (error) {
        stats.engineErrors++;
        throw error instanceof Error ? error : new Error(String(error));
      }
      stats.rustPaginated++;
      return layout;
    },
  };
}
