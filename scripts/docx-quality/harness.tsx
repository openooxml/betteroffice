import React from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor } from '@betteroffice/docx-react';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import '@betteroffice/docx-react/styles.css';
import { createFontProvider } from '../../packages/fonts/src/cdn';
import { MAX_LOCAL_REFERENCE_PAGES } from '../office-quality/reference.mjs';
import {
  capturePageExtent,
  extentMatches,
  validatePageBounds,
  type OfficePageBounds,
} from '../office-quality/page-bounds';
import { hasInk, isUnpaintedCanvas, layoutWaitMs, settleCanvas } from './capture-wait';

declare const __QUALITY_FONT_BASE__: string | null;
declare const __QUALITY_FONT_BASE_CJK__: string | null;

const search = new URLSearchParams(window.location.search);
// Local-only captures load this checkout's fonts from the viewer's own origin.
const localFonts = search.get('local') === '1' && typeof __QUALITY_FONT_BASE__ === 'string';
const legacySyncApi = search.get('peerless') === '0';
const provider = createFontProvider(
  localFonts
    ? {
        baseUrl: __QUALITY_FONT_BASE__,
        ...(typeof __QUALITY_FONT_BASE_CJK__ === 'string'
          ? { cjkBaseUrl: __QUALITY_FONT_BASE_CJK__ }
          : {}),
      }
    : undefined
);
setGoogleFontsEnabled(false);
const fontLoads: {
  kind: string;
  family: string;
  ok: boolean;
  error?: string;
}[] = [];
const track =
  (kind: string, resolve: (...args: any[]) => (() => Promise<ArrayBuffer>) | undefined) =>
  (...args: any[]) => {
    const load = resolve(...args);
    if (!load) return undefined;
    return async () => {
      try {
        const bytes = await load();
        fontLoads.push({ kind, family: String(args[0]), ok: true });
        return bytes;
      } catch (error) {
        fontLoads.push({
          kind,
          family: String(args[0]),
          ok: false,
          error: String(error),
        });
        throw error;
      }
    };
  };
const measuredFonts = {
  resolve: track('family', provider.resolve),
  resolveScriptFallback: track('script', provider.resolveScriptFallback),
  resolveLastResort: track('fallback', provider.resolveLastResort),
};
let editor: React.RefObject<any>;
let root: ReturnType<typeof createRoot>;
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
const api = window as any;
let captureProfile: OfficePageBounds | null = null;
const pageExtents: ReturnType<typeof capturePageExtent>[] = [];
const requestedMaxPages = search.get('maxPages');
const maxPages = requestedMaxPages === null ? undefined : Number(requestedMaxPages);
if (maxPages !== undefined && (!Number.isInteger(maxPages) || maxPages < 1))
  throw new Error('maxPages must be a positive integer');
api.oracleInit = async (input: number[], fonts: boolean, profile?: unknown) => {
  captureProfile = validatePageBounds(
    profile,
    maxPages === undefined ? undefined : Math.min(maxPages, MAX_LOCAL_REFERENCE_PAGES)
  );
  pageExtents.length = 0;
  fontLoads.length = 0;
  const bytes = new Uint8Array(input).buffer;
  root?.unmount();
  editor = React.createRef();
  api.errors = [];
  root = createRoot(document.getElementById('host')!);
  api.stage = 'editor';
  root.render(
    <DocxEditor
      ref={editor}
      {...(legacySyncApi ? { legacySyncApi: true as const } : {})}
      documentBuffer={bytes}
      readOnly
      initialZoom={1}
      colorMode="light"
      showZoomControl={false}
      showOutlineButton={false}
      measurementFontProvider={fonts ? measuredFonts : undefined}
      onError={(error) => api.errors.push(String(error))}
    />
  );
  let last = 0;
  let stable = 0;
  // Releases without whenLayoutComplete lay the whole document out before the first paint.
  let complete: boolean | null = null;
  let failure: unknown = null;
  for (;;) {
    if (api.errors.length > 0) throw new Error(api.errors.join('\n'));
    if (failure) throw failure;
    if (complete === null && editor.current) {
      complete = typeof editor.current.whenLayoutComplete !== 'function';
      if (!complete) {
        editor.current.whenLayoutComplete({ timeoutMs: layoutWaitMs(captureProfile?.pages.length) }).then(
          () => (complete = true),
          (error: unknown) => (failure = error)
        );
      }
    }
    const pages = complete ? editor.current.getTotalPages() : 0;
    const canvas = document.querySelector<HTMLCanvasElement>(
      'canvas[data-page-index="0"]'
    );
    if (pages > 0 && pages === last && canvas?.width && canvas?.height) {
      if (Date.now() - stable >= 3000) {
        api.stage = 'ready';
        return { pages, errors: api.errors, fontLoads };
      }
    } else {
      last = pages;
      stable = Date.now();
    }
    await sleep(100);
  }
};
const canvasHasInk = (canvas: HTMLCanvasElement) => {
  const probe = document.createElement('canvas');
  probe.width = canvas.width;
  probe.height = canvas.height;
  const context = probe.getContext('2d', { willReadFrequently: true })!;
  context.drawImage(canvas, 0, 0);
  const inked = hasInk(context.getImageData(0, 0, probe.width, probe.height).data);
  probe.width = 0;
  probe.height = 0;
  return inked;
};
api.oraclePage = async (index: number) => {
  editor.current.scrollToPage(index + 1);
  const frame = await settleCanvas({
    read: () => {
      const canvas = document.querySelector<HTMLCanvasElement>(
        `canvas[data-page-index="${index}"]`
      );
      if (!canvas?.width || !canvas?.height) return null;
      const extent = capturePageExtent(captureProfile, index, canvas.width, canvas.height, true);
      const output = document.createElement('canvas');
      output.width = extent.output.width_px;
      output.height = extent.output.height_px;
      const context = output.getContext('2d')!;
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, output.width, output.height);
      context.drawImage(canvas, 0, 0);
      const value = output.toDataURL('image/png');
      output.width = 0;
      output.height = 0;
      const painted =
        !isUnpaintedCanvas(canvas.width, canvas.height, false) || canvasHasInk(canvas);
      return { width: canvas.width, height: canvas.height, value, painted };
    },
    matches: (width, height) => extentMatches(captureProfile, index, width, height),
    sleep,
    now: () => Date.now(),
  });
  // local captures keep a page whose size differs from Word's page at that index
  pageExtents[index] = capturePageExtent(
    captureProfile,
    index,
    frame.width,
    frame.height,
    maxPages !== undefined
  );
  return frame.value;
};
api.oracleCaptureMetadata = () => ({
  capture_profile: captureProfile,
  page_extents: pageExtents,
});
api.oracleReady = true;
