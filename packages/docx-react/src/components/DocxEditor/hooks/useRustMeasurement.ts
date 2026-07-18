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

export function useRustMeasurement(
  options: UseRustMeasurementOptions
): UseRustMeasurementReturn {
  const { document, fontProvider, fontChainsProviderRef, textEngine } = options;
  const runLayoutPipelineRef = useRef<(() => void) | null>(null);
  const sourceRef = useRef<RustMeasureSource | null>(null);
  const sourceEngineRef = useRef<RustTextEngine | null>(null);
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
          source = createRustMeasureSource({ engine, bundled: fontProviderRef.current });
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
          latestFontChainsRef.current = {};
        }
        if (firstLoad) runLayoutPipelineRef.current?.();
      } catch (error) {
        console.error('[useRustMeasurement] Rust font engine failed to load', error);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [document, textEngine]);

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
        const warmup = source
          .prepareFontRequirements(requirements)
          .then(() => {
            if (sourceRef.current === source) runLayoutPipelineRef.current?.();
          })
          .catch(() => {})
          .finally(() => requirementWarmupsRef.current.delete(key));
        requirementWarmupsRef.current.set(key, warmup);
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
