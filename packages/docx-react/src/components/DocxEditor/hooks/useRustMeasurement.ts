/**
 * Host wiring for the Rust measurement source — the sole measurement path.
 *
 * The hook lazily loads the wasm text engine, builds a `RustMeasureSource`
 * over it, feeds it the open document's embedded faces and compat flags, and
 * warms the font registry for the current block list (`prepareFonts` — the
 * only async step).
 *
 * Readiness gate: the layout pipeline calls `deferLayoutPass()` before
 * committing a pass. While the engine is still loading, or while a font
 * chain some measured paragraph needed is still resolving, the pass is
 * discarded and this hook re-runs the pipeline (through
 * `runLayoutPipelineRef`) once the engine/fonts settle — so no committed
 * layout is ever measured against an unloaded face.
 */

import { useCallback, useEffect, useRef } from 'react';

import type { Document, FontTable } from '@betteroffice/docx/types/document';
import type { BlockExtent, LayoutBlock } from '@betteroffice/docx/layout/pagination';
import {
  createRustMeasureSource,
  getRustTextEngine,
  measureBlocksWithFloats,
  type BundledFontProvider,
  type FloatPageGeometry,
  type RustMeasureSource,
  type RustTextEngine,
} from '@betteroffice/docx/layout';
import { extractEmbeddedFontFaces } from '@betteroffice/docx/utils';


/**
 * Debug flag: `?rustMeasureDebug=1` logs a one-line rust-vs-synthetic block
 * count after every layout pass. Read once at module load.
 */
const DEBUG_STATS =
  typeof window !== 'undefined' && window.location.search.includes('rustMeasureDebug=1');

/**
 * Reads the merged doc-wide font-id chain map for the latest block list, or
 * `undefined` when no chains have resolved yet. The canvas display-list build
 * calls this at build time to gate GlyphRun emission.
 */
export type RustFontChainsProvider = () => Record<string, number[]> | undefined;

export interface UseRustMeasurementOptions {
  /** Current document — source of embedded faces and compat flags. */
  document: Document | null;
  /** Optional bundled metric-compatible font provider (`measurementFontProvider`). */
  fontProvider?: BundledFontProvider;
  /**
   * Host-owned slot the hook fills with a {@link RustFontChainsProvider} so the
   * canvas display-list build (mounted in a parent) can read the merged font
   * chains at build time. Cleared to `null` on unmount. Omit when there is no
   * canvas renderer to feed.
   */
  fontChainsProviderRef?: React.RefObject<RustFontChainsProvider | null>;
  /** Resident editing-wasm engine; null uses the standalone layout module. */
  textEngine?: RustTextEngine | null;
}

export interface UseRustMeasurementReturn {
  /**
   * The `measureBlocks` implementation for the layout pipeline. Must only be
   * called on a pass `deferLayoutPass()` did not defer (it throws if the
   * engine has not loaded yet).
   */
  measureBlocksImpl: (
    blocks: LayoutBlock[],
    contentWidth: number | number[],
    pageGeometry?: FloatPageGeometry
  ) => BlockExtent[];
  /** Feed each layout pass's blocks so font preparation tracks the document. */
  notifyBlocks: (blocks: LayoutBlock[]) => void;
  /**
   * The readiness gate. True while the engine is loading or a font chain a
   * measure pass needed is still resolving — the pipeline must skip/discard
   * the pass; this hook re-runs it once the engine/fonts settle. Passing the
   * pass's blocks lets the settle-warmup cover them too.
   */
  deferLayoutPass: (blocks?: LayoutBlock[]) => boolean;
  /**
   * The host assigns `runLayoutPipeline` here right after useLayoutPipeline
   * returns. The hook re-runs the pipeline through it when the engine
   * becomes ready and when newly resolved fonts settle.
   */
  runLayoutPipelineRef: React.RefObject<(() => void) | null>;
}

export function useRustMeasurement(opts: UseRustMeasurementOptions): UseRustMeasurementReturn {
  const { document, fontProvider, fontChainsProviderRef, textEngine } = opts;
  const runLayoutPipelineRef = useRef<(() => void) | null>(null);

  const sourceRef = useRef<RustMeasureSource | null>(null);
  const sourceEngineRef = useRef<RustTextEngine | null>(null);
  const latestBlocksRef = useRef<LayoutBlock[]>([]);
  // Which (buffer, fontTable) pair the source's embedded faces came from.
  // The yrs projection carries both by reference across edits, so this only
  // changes on a real document load. The guard keeps per-keystroke effect
  // runs from re-unzipping the package and invalidating the measure memo.
  const fedFontSourceRef = useRef<{
    buffer: ArrayBuffer | null;
    fontTable: FontTable | null;
  } | null>(null);
  const lastLoggedStatsRef = useRef({ rustMeasured: 0, syntheticFallback: 0 });

  // Latest-value ref so the stable callbacks below never go stale.
  const fontProviderRef = useRef(fontProvider);
  fontProviderRef.current = fontProvider;

  const prepare = useCallback(
    (blocks: LayoutBlock[]) => {
      const source = sourceRef.current;
      if (!source) return;
      void source
        .prepareFonts(blocks)
        .then((newlyAvailable) => {
          if (!newlyAvailable) return;
          if (sourceRef.current !== source) return;
          runLayoutPipelineRef.current?.();
        })
        .catch(() => {
          // prepareFonts is failure-tolerant by contract; a rejection here
          // means those chains behave like settled misses.
        });
    },
    []
  );

  // Engine + source lifecycle, and per-document sync (compat flags,
  // embedded faces). Runs on every `document` identity change (each edit
  // pushes a new Document), so everything repeated here is either O(1)
  // (setCompat, prepare over settled chains) or guarded (face extraction).
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const engine = textEngine ?? (await getRustTextEngine());
        if (cancelled) return;
        let source = sourceRef.current;
        if (sourceEngineRef.current !== engine) {
          source = null;
          sourceRef.current = null;
          sourceEngineRef.current = engine;
          fedFontSourceRef.current = null;
        }
        const firstLoad = !source;
        if (!source) {
          source = createRustMeasureSource({
            engine,
            // Deterministic resolution: embedded document faces + the injected
            // bundled provider only (no OS/local-font source).
            bundled: fontProviderRef.current,
          });
          sourceRef.current = source;
        }
        source.setCompat(document?.package.settings?.compatibilityFlags);

        const buffer = document?.originalBuffer ?? null;
        const fontTable = document?.package.fontTable ?? null;
        const fed = fedFontSourceRef.current;
        if (!fed || fed.buffer !== buffer || fed.fontTable !== fontTable) {
          const faces = document ? await extractEmbeddedFontFaces(document) : [];
          if (cancelled) return;
          source.setEmbeddedFaces(faces);
          fedFontSourceRef.current = { buffer, fontTable };
        }

        if (firstLoad) {
          // Un-defer the passes that ran before the engine loaded — this
          // first re-run measures, records any unresolved chains, and defers
          // again until prepareFonts settles them.
          runLayoutPipelineRef.current?.();
        }

        prepare(latestBlocksRef.current);
      } catch (error) {
        console.error('[useRustMeasurement] Rust text engine failed to load — no layout', error);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [document, prepare, textEngine]);

  const notifyBlocks = useCallback(
    (blocks: LayoutBlock[]) => {
      latestBlocksRef.current = blocks;
      prepare(blocks);
      if (DEBUG_STATS && sourceRef.current) {
        const stats = sourceRef.current.getStats();
        const prev = lastLoggedStatsRef.current;
        if (
          stats.rustMeasured !== prev.rustMeasured ||
          stats.syntheticFallback !== prev.syntheticFallback
        ) {
          // eslint-disable-next-line no-console -- opt-in debug output (?rustMeasureDebug=1)
          console.debug(
            `[useRustMeasurement] pass: rust ${stats.rustMeasured - prev.rustMeasured}, ` +
              `synthetic ${stats.syntheticFallback - prev.syntheticFallback} ` +
              `(cumulative ${stats.rustMeasured}/${stats.syntheticFallback}; ` +
              `font-unready ${stats.syntheticFontUnready}, ` +
              `uncovered ${stats.syntheticUncovered})`
          );
          lastLoggedStatsRef.current = stats;
        }
      }
    },
    [prepare]
  );

  // The readiness gate — see UseRustMeasurementReturn.deferLayoutPass.
  const deferLayoutPass = useCallback(
    (blocks?: LayoutBlock[]): boolean => {
      const source = sourceRef.current;
      // Engine still loading: the lifecycle effect re-runs the pipeline the
      // moment the source exists.
      if (!source) return true;
      if (!source.hasPendingFonts()) return false;
      // Chains a measure pass needed are still resolving — warm them
      // (idempotent) and re-run when they settle.
      prepare(blocks ?? latestBlocksRef.current);
      return true;
    },
    [prepare]
  );

  // Stable identity — reads the source through a ref so useLayoutPipeline's
  // dep array (and its rAF scheduler) never churns.
  const measureBlocksImpl = useCallback(
    (
      blocks: LayoutBlock[],
      contentWidth: number | number[],
      pageGeometry?: FloatPageGeometry
    ): BlockExtent[] => {
      const source = sourceRef.current;
      if (!source) {
        throw new Error(
          '[useRustMeasurement] measure pass ran before the engine was ready — ' +
            'the layout pipeline must defer via deferLayoutPass()'
        );
      }
      return measureBlocksWithFloats(
        blocks,
        contentWidth,
        source.createMeasureBlock(),
        pageGeometry
      );
    },
    []
  );

  // Merged doc-wide font chains for the latest pass's blocks — the map the
  // canvas display-list build consumes to gate GlyphRun emission. Stable
  // identity (reads through refs); returns undefined before the engine loads
  // or before any chain resolves (⇒ the build omits fontChains and keeps the
  // char-distributed draw path).
  const getDocumentFontChains = useCallback<RustFontChainsProvider>(() => {
    const source = sourceRef.current;
    if (!source) return undefined;
    const chains = source.getDocumentFontChains(latestBlocksRef.current);
    return Object.keys(chains).length > 0 ? chains : undefined;
  }, []);

  // Fill the host-owned slot so the parent's display-list build can read the
  // chains at build time; clear it on unmount so a stale getter never lingers.
  useEffect(() => {
    if (!fontChainsProviderRef) return;
    fontChainsProviderRef.current = getDocumentFontChains;
    return () => {
      if (fontChainsProviderRef.current === getDocumentFontChains) {
        fontChainsProviderRef.current = null;
      }
    };
  }, [fontChainsProviderRef, getDocumentFontChains]);

  return {
    measureBlocksImpl,
    notifyBlocks,
    deferLayoutPass,
    runLayoutPipelineRef,
  };
}
