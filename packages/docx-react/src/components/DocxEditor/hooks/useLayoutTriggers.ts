/**
 * Layout-trigger effects for PagedEditor.
 *
 * Re-runs `runLayoutPipeline` for two state-shifts that the pipeline's
 * own dep array doesn't catch automatically:
 *
 *  1. Web-font loading completes — measurement is Rust-only (font bytes,
 *     not browser fonts), so the layout itself cannot change; the re-run
 *     re-rasterizes canvas text drawn through CSS-font fallback paths so
 *     late-loading embedded faces show up.
 *
 *  2. Header / footer content or render-env changes — runLayoutPipeline does include
 *     these in its deps, but only re-runs when explicitly called. The
 *     first render already laid out when the Yrs session became ready, so this
 *     effect skips the initial render via a one-shot epoch counter.
 */

import { useEffect, useRef } from 'react';

import type { HeaderFooter } from '@betteroffice/docx/types/document';
import type { YrsRenderEnv } from '@betteroffice/docx/yrs';
export interface UseLayoutTriggersOptions {
  runLayoutPipeline: () => void;
  updateSelectionOverlay: () => void;
  headerContent?: HeaderFooter | null;
  footerContent?: HeaderFooter | null;
  firstPageHeaderContent?: HeaderFooter | null;
  firstPageFooterContent?: HeaderFooter | null;
  renderEnv?: YrsRenderEnv;
}

export function useLayoutTriggers(opts: UseLayoutTriggersOptions): void {
  const {
    runLayoutPipeline,
    updateSelectionOverlay,
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    renderEnv,
  } = opts;
  const runLayoutPipelineRef = useRef(runLayoutPipeline);
  runLayoutPipelineRef.current = runLayoutPipeline;
  const updateSelectionOverlayRef = useRef(updateSelectionOverlay);
  updateSelectionOverlayRef.current = updateSelectionOverlay;

  // Re-layout on web-font load. FontFaceSet.onloadingdone catches new
  // fonts as they finish loading.
  useEffect(() => {
    const handleFontsLoaded = () => {
      runLayoutPipelineRef.current();
      updateSelectionOverlayRef.current();
    };
    window.document.fonts.addEventListener('loadingdone', handleFontsLoaded);
    return () => {
      window.document.fonts.removeEventListener('loadingdone', handleFontsLoaded);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Re-layout when H/F content or the render env changes (HF editor save,
  // showHiddenText toggle, etc.).
  const contentEpochRef = useRef(0);
  useEffect(() => {
    // Skip the initial render — session readiness already triggered the first layout.
    if (contentEpochRef.current === 0) {
      contentEpochRef.current = 1;
      return;
    }
    runLayoutPipelineRef.current();
  }, [
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    renderEnv,
  ]);
}
