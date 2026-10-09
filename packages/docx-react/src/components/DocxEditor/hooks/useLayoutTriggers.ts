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
import { yieldToMainThread } from '../internals/yieldToMainThread';
export interface UseLayoutTriggersOptions {
  runLayoutPipeline: () => void;
  updateSelectionOverlay: () => void;
  headerContent?: HeaderFooter | null;
  footerContent?: HeaderFooter | null;
  firstPageHeaderContent?: HeaderFooter | null;
  firstPageFooterContent?: HeaderFooter | null;
  renderEnv?: YrsRenderEnv;
  holdFontRefresh?: () => boolean;
  fontRefreshReleased?: boolean;
  fontRefreshScope?: unknown;
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
    holdFontRefresh,
    fontRefreshReleased,
    fontRefreshScope,
  } = opts;
  const runLayoutPipelineRef = useRef(runLayoutPipeline);
  runLayoutPipelineRef.current = runLayoutPipeline;
  const updateSelectionOverlayRef = useRef(updateSelectionOverlay);
  updateSelectionOverlayRef.current = updateSelectionOverlay;
  const holdFontRefreshRef = useRef(holdFontRefresh);
  holdFontRefreshRef.current = holdFontRefresh;
  const fontRefreshReleasedRef = useRef(fontRefreshReleased);
  fontRefreshReleasedRef.current = fontRefreshReleased;
  const committedFontRefreshScopeRef = useRef(fontRefreshScope);
  const heldFontRefreshRef = useRef<{ scope: unknown } | null>(null);

  useEffect(() => {
    committedFontRefreshScopeRef.current = fontRefreshScope;
    const held = heldFontRefreshRef.current;
    if (held && !Object.is(held.scope, fontRefreshScope)) {
      heldFontRefreshRef.current = null;
    }
  }, [fontRefreshScope]);

  // Re-layout on web-font load. FontFaceSet.onloadingdone catches new
  // fonts as they finish loading.
  useEffect(() => {
    const handleFontsLoaded = () => {
      const scope = committedFontRefreshScopeRef.current;
      const held = heldFontRefreshRef.current;
      if ((held && Object.is(held.scope, scope)) || holdFontRefreshRef.current?.()) {
        heldFontRefreshRef.current = { scope };
        return;
      }
      runLayoutPipelineRef.current();
      updateSelectionOverlayRef.current();
    };
    window.document.fonts.addEventListener('loadingdone', handleFontsLoaded);
    return () => {
      heldFontRefreshRef.current = null;
      window.document.fonts.removeEventListener('loadingdone', handleFontsLoaded);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    const held = heldFontRefreshRef.current;
    if (!fontRefreshReleased || !held || !Object.is(held.scope, fontRefreshScope)) return;
    let cancelled = false;
    const replay = async () => {
      await yieldToMainThread();
      if (
        cancelled ||
        !fontRefreshReleasedRef.current ||
        !heldFontRefreshRef.current ||
        !Object.is(heldFontRefreshRef.current.scope, fontRefreshScope) ||
        !Object.is(committedFontRefreshScopeRef.current, fontRefreshScope)
      ) return;
      heldFontRefreshRef.current = null;
      runLayoutPipelineRef.current();
      updateSelectionOverlayRef.current();
    };
    void replay();
    return () => { cancelled = true; };
  }, [fontRefreshReleased, fontRefreshScope]);

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
