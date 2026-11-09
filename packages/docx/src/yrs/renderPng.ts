import { buildResidentRegionLayoutRequest, computeLayout, documentRenderEnv, getLayoutKernelInputs } from '../editor/computeLayout';
import { createRustMeasureSource, resolveCompleteFontProvider, type BundledFontProvider, type ResidentFontRequirement } from '../layout/measure';
import type { EmbeddedFaceInput } from '../layout/measure/fontRegistry';
import { fontUnavailable, pngPage, pngScale, type DocxPngFontFace, type DocxPngOptions, type DocxPngTarget } from '../layout/render/png';
import { buildRustDisplayList } from '../layout/render/rustDisplayList';
import type { DisplayList } from '../layout/render/displayList';
import { extractEmbeddedFontFaces } from '../utils';
import { engineKind, renderPng as renderEnginePng } from '../wasm/enginePng';
import { createYrsSession, type YrsSession } from './index';
import { sessionMediaSources } from './sessionInternals';

const IDLE_RELEASE_MS = 30_000;

interface Held {
  version: string;
  fonts: ReadonlyArray<DocxPngFontFace> | undefined;
  fontProvider: BundledFontProvider | undefined;
  laidOut: Promise<{ copy: YrsSession; displayList: DisplayList; fontChains: Record<string, number[]> }>;
  users: number;
  stale: boolean;
}

function appFaces(fonts: ReadonlyArray<DocxPngFontFace> | undefined): EmbeddedFaceInput[] {
  return (fonts ?? []).map((face) => ({
    family: face.family, weight: face.bold ? 'bold' : 'normal', style: face.italic ? 'italic' : 'normal',
    data: face.bytes.slice().buffer as ArrayBuffer, subsetted: false,
  }));
}

/**
 * Lays a replica of `session` out once per version with the editor's chain and
 * draws its pages. The replica opens from the session's own state, so it holds
 * the model the session has, and the session's retained layout stays untouched.
 */
export function createPngExporter(session: YrsSession) {
  let held: Held | undefined;
  let idle: ReturnType<typeof setTimeout> | undefined;
  let alive = true;

  const assertAlive = () => { if (!alive) throw new Error('YrsSession is destroyed'); };
  const free = (entry: Held) => { void entry.laidOut.then(({ copy }) => copy.destroy(), () => {}); };
  const release = () => {
    if (idle !== undefined) clearTimeout(idle);
    idle = undefined;
    if (!held) return;
    held.stale = true;
    if (held.users === 0) free(held);
    held = undefined;
  };

  async function layOut(state: Uint8Array, metadata: Uint8Array, media: string, source: ArrayBuffer, fonts: Held['fonts'], fontProvider: Held['fontProvider']): Promise<Held['laidOut'] extends Promise<infer T> ? T : never> {
    const copy = await createYrsSession();
    try {
      const bytes = new Uint8Array(source);
      const host = copy.openDocx(bytes, false, { mediaTokens: true });
      copy.bootstrapPeer(state, metadata, bytes, host, { mediaTokens: true });
      if (media) copy.loadMediaSources(media);
      const document = copy.materializeDocx();
      if (!document) throw new Error('renderPng needs a session opened from DOCX bytes');
      const bundled = fontProvider ?? (engineKind === 'native' ? await resolveCompleteFontProvider() : undefined);
      const measure = createRustMeasureSource({ engine: copy, ...(bundled ? { bundled } : { standIns: true }) });
      measure.setCompat(document.package.settings?.compatibilityFlags);
      measure.setEmbeddedFaces(await extractEmbeddedFontFaces(document), appFaces(fonts));
      const renderEnv = documentRenderEnv(document);
      const request = buildResidentRegionLayoutRequest(document, 0, renderEnv);
      const requirements = JSON.parse(copy.layoutFontRequirementsJson(JSON.stringify(request))) as ResidentFontRequirement[];
      await measure.prepareFontRequirements(requirements);
      const measurement = measure.measurementConfigForRequirements(requirements);
      if (!measurement) throw fontUnavailable('The fonts this document lays out with could not be loaded. Install @betteroffice/fonts, or pass them in `fonts`.');
      const { layout } = computeLayout({ document, pageGap: 0, session: copy, renderEnv, measurement });
      const displayList = await buildRustDisplayList({ ...getLayoutKernelInputs(layout)!, layout, fontChains: measurement.fontChains }, copy);
      return { copy, displayList, fontChains: measurement.fontChains ?? {} };
    } catch (error) {
      try { copy.destroy(); } catch { /* the layout failure is the one to report */ }
      throw error;
    }
  }

  function current(fonts: Held['fonts'], fontProvider: Held['fontProvider']): Held {
    const version = session.version();
    if (held && held.version === version && held.fonts === fonts && held.fontProvider === fontProvider) return held;
    release();
    const source = session.materializeDocx()?.originalBuffer;
    if (!source) throw new Error('renderPng needs a session opened from DOCX bytes');
    const entry: Held = {
      version, fonts, fontProvider, users: 0, stale: false,
      laidOut: layOut(session.encodeState(), session.encodePeerMetadata(), sessionMediaSources(session), source, fonts, fontProvider),
    };
    entry.laidOut.catch(() => { if (held === entry) held = undefined; });
    held = entry;
    return entry;
  }

  return {
    async render(target: DocxPngTarget, options: DocxPngOptions = {}): Promise<Uint8Array> {
      assertAlive();
      const index = pngPage(target);
      pngScale(options.scale);
      options.signal?.throwIfAborted();
      const entry = current(options.fonts, options.fontProvider);
      entry.users++;
      if (idle !== undefined) clearTimeout(idle);
      idle = undefined;
      try {
        const { copy, displayList, fontChains } = await entry.laidOut;
        assertAlive();
        const page = displayList.pages[index];
        if (!page) throw new RangeError(`page index ${index} out of range for ${displayList.pages.length} page(s)`);
        const bytes = await renderEnginePng({
          displayList, fontChains, page, session: copy,
          outlineGlyph: (fontId, glyphId) => copy.outlineGlyphJson(fontId, glyphId),
          media: (token) => copy.mediaSource(token),
        }, options);
        assertAlive();
        return bytes;
      } finally {
        entry.users--;
        if (entry.stale && entry.users === 0) free(entry);
        if (held === entry && entry.users === 0 && alive) {
          idle = setTimeout(release, IDLE_RELEASE_MS);
          (idle as { unref?: () => void }).unref?.();
        }
      }
    },
    dispose() { alive = false; release(); },
  };
}

