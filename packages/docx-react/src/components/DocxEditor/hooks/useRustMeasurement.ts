import { useCallback, useEffect, useRef } from 'react';

import type { Document, FontTable } from '@betteroffice/docx/types/document';
import {
  createRustMeasureSource,
  getRustTextEngine,
  type BundledFontProvider,
  type ResidentFontRequirement,
  type ResidentMeasurementConfig,
  type RustMeasureSource,
  type RustTextEngine,
} from '@betteroffice/docx/layout';
import { extractEmbeddedFontFaces } from '@betteroffice/docx/utils';

export type RustFontChainsProvider = () => Record<string, number[]> | undefined;

export interface UseRustMeasurementOptions {
  /** `textEngine`: the engine whose fonts failed to load, if any. */
  onError?: (error: Error, textEngine?: RustTextEngine | null) => void;
  document: Document | null;
  fontProvider?: BundledFontProvider;
  fontChainsProviderRef?: React.RefObject<RustFontChainsProvider | null>;
  textEngine?: RustTextEngine | null;
}

export interface UseRustMeasurementReturn {
  deferLayoutPass: () => boolean;
  residentMeasurementConfig: (
    requirements: ResidentFontRequirement[]
  ) => ResidentMeasurementConfig | null;
  runLayoutPipelineRef: React.RefObject<(() => void) | null>;
}

/** `engine` until `release()`: fonts that finish loading after it register nothing. */
function releasableTextEngine(engine: RustTextEngine): {
  engine: RustTextEngine;
  release: () => void;
} {
  let released = false;
  return {
    engine: {
      registerFont: (bytes) => (released ? -1 : engine.registerFont(bytes)),
      ...(engine.registerSubstituteFont && {
        registerSubstituteFont: (id: number, family: string) =>
          released ? id : engine.registerSubstituteFont!(id, family),
      }),
      clearFonts: () => {
        if (!released) engine.clearFonts();
      },
    },
    release: () => {
      released = true;
    },
  };
}

export function useRustMeasurement(
  options: UseRustMeasurementOptions
): UseRustMeasurementReturn {
  const { document, fontProvider, fontChainsProviderRef, textEngine } = options;
  const onErrorRef = useRef(options.onError);
  onErrorRef.current = options.onError;
  const runLayoutPipelineRef = useRef<(() => void) | null>(null);
  const sourceRef = useRef<RustMeasureSource | null>(null);
  const sourceEngineRef = useRef<RustTextEngine | null>(null);
  const releaseSourceRef = useRef<(() => void) | null>(null);
  const latestFontChainsRef = useRef<Record<string, number[]>>({});
  const requirementWarmupsRef = useRef(new Map<string, Promise<void>>());
  const fedFontSourceRef = useRef<{
    buffer: ArrayBuffer | null;
    fontTable: FontTable | null;
  } | null>(null);
  const fontProviderRef = useRef(fontProvider);
  fontProviderRef.current = fontProvider;

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
          latestFontChainsRef.current = {};
          requirementWarmupsRef.current.clear();
        }
        const firstLoad = !source;
        if (!source) {
          const releasable = releasableTextEngine(engine);
          source = createRustMeasureSource({
            engine: releasable.engine,
            bundled: fontProviderRef.current,
          });
          sourceRef.current = source;
          releaseSourceRef.current = releasable.release;
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
          latestFontChainsRef.current = {};
        }
        if (firstLoad) runLayoutPipelineRef.current?.();
      } catch (error) {
        console.error('[useRustMeasurement] Rust font engine failed to load', error);
        if (!cancelled) {
          onErrorRef.current?.(error instanceof Error ? error : new Error(String(error)), textEngine);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [document, textEngine]);

  // A replaced or unmounted editor may free the session behind `textEngine`: font loads and
  // layout passes its source still has pending end with it.
  useEffect(
    () => () => {
      releaseSourceRef.current?.();
      releaseSourceRef.current = null;
      sourceRef.current = null;
      // The next source starts over, even on the same engine.
      sourceEngineRef.current = null;
    },
    [textEngine]
  );

  const deferLayoutPass = useCallback((): boolean => sourceRef.current === null, []);

  const residentMeasurementConfig = useCallback(
    (requirements: ResidentFontRequirement[]): ResidentMeasurementConfig | null => {
      const source = sourceRef.current;
      if (!source) return null;
      const ready = source.measurementConfigForRequirements(requirements);
      if (ready) {
        latestFontChainsRef.current = ready.fontChains;
        return ready;
      }
      const key = JSON.stringify(requirements);
      if (!requirementWarmupsRef.current.has(key)) {
        const settled = source
          .prepareFontRequirements(requirements)
          .then(
            () => undefined,
            () => undefined
          )
          .finally(() => {
            if (sourceRef.current !== source) return;
            requirementWarmupsRef.current.delete(key);
            runLayoutPipelineRef.current?.();
          });
        requirementWarmupsRef.current.set(key, settled);
      }
      return null;
    },
    []
  );

  const getDocumentFontChains = useCallback<RustFontChainsProvider>(() => {
    const chains = latestFontChainsRef.current;
    return Object.keys(chains).length > 0 ? chains : undefined;
  }, []);

  useEffect(() => {
    if (!fontChainsProviderRef) return;
    fontChainsProviderRef.current = getDocumentFontChains;
    return () => {
      if (fontChainsProviderRef.current === getDocumentFontChains) {
        fontChainsProviderRef.current = null;
      }
    };
  }, [fontChainsProviderRef, getDocumentFontChains]);

  return { deferLayoutPass, residentMeasurementConfig, runLayoutPipelineRef };
}
