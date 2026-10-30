import { createPptxWorkerSession, decodeCommandDisplayList, browserFrameImages } from '@betteroffice/pptx/worker';
import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx';
import { installFonts } from '../../packages/pptx-react/src/fontFaces';
import { installDefaultFonts } from '../../packages/pptx-react/src/worker/peerlessFonts';
import { BUNDLED_FONTS, WORD_FAMILY_ALIASES } from '../../packages/fonts/src/manifest';
import { loadFontBytes } from '../../packages/fonts/src/bytes';
import { loadPackedFontBytes } from '../../packages/fonts/src/packed';

declare const __QUALITY_FONT_BASE__: string | null;
declare const __QUALITY_FONT_BASE_CJK__: string | null;

type Face = { family: string; bold: boolean; italic: boolean; bytes: Uint8Array };
type Provider = {
  resolve(family: string, bold: boolean, italic: boolean): (() => Promise<ArrayBuffer>) | undefined;
  resolveLastResort(family: string, bold: boolean, italic: boolean, office?: string): () => Promise<ArrayBuffer>;
};

async function hostFonts(host: string, provider: Provider): Promise<Face[]> {
  const fonts: Face[] = [];
  if (host === 'arial') {
    for (const [bold, italic] of [[false, false], [true, false], [false, true], [true, true]]) {
      const load = provider.resolve('Arial', bold, italic) ?? provider.resolveLastResort('Arial', bold, italic, 'powerpoint');
      fonts.push({ family: 'Arial', bold, italic, bytes: new Uint8Array(await load()) });
    }
  } else if (host === 'bundled') {
    for (const face of BUNDLED_FONTS) {
      if (typeof __QUALITY_FONT_BASE__ !== 'string' || typeof __QUALITY_FONT_BASE_CJK__ !== 'string')
        throw new Error('The bundled host needs local font assets');
      const base = new URL(__QUALITY_FONT_BASE__, location.origin);
      const bytes = new Uint8Array(face.file.endsWith('.ttf.zst')
        ? await loadPackedFontBytes(face, async (dependency) => new URL(dependency.file, base), 30_000,
          new URL('decoder/fontpack.wasm', base))
        : await loadFontBytes(face,
          new URL(face.file, face.script?.startsWith('cjk-') ? new URL(__QUALITY_FONT_BASE_CJK__, location.origin) : base), 30_000));
      const names = new Set([face.family]);
      if (face.metricCompatWith) names.add(face.metricCompatWith);
      const targets = new Set([...names].map((name) => name.toLowerCase()));
      for (const [alias, target] of Object.entries(WORD_FAMILY_ALIASES))
        if (targets.has(target.toLowerCase())) names.add(alias);
      if (face.family === 'Carlito') names.add('Calibri Light');
      for (const family of names) fonts.push({ family, bold: face.weight === 700, italic: face.style === 'italic', bytes });
    }
  } else if (host !== 'none') throw new Error(`Unknown host fonts ${host}`);
  return fonts;
}

/**
 * Opens a deck the way the default editor does: the command worker, app fonts
 * installed in the page and default faces installed through the editor's own
 * installer. `host` is what the app supplies: `arial` (four Arial faces),
 * `bundled` (every bundled face under its family, metric-compatible name and
 * Word aliases) or `none`.
 */
export async function openWorkerCapture(bytes: Uint8Array, host: string, provider: Provider, defaultFonts: boolean) {
  const fonts = await hostFonts(host, provider);
  const installed: FontFace[] = [];
  const supplied = new Set(fonts.map((face) => face.family.toLowerCase()));
  await installFonts(fonts as never, installed, () => true);
  const session = createPptxWorkerSession(bytes, Object.assign({ fonts, defaultFonts }, {
    pptxDefaultFontInstaller: (faces: Parameters<typeof installDefaultFonts>[0], allowed?: (family: string) => boolean) =>
      installDefaultFonts(faces, installed, () => true, () => {},
        (family) => !supplied.has(family.trim().toLowerCase()) && (allowed?.(family) ?? true)),
  }) as never);
  const client: any = await session.ready;
  const slides = client.state.slides as Array<{ id: string }>;
  const images = browserFrameImages();
  // Media bytes arrive once per asset, so later thumbnails reference assets earlier frames carried.
  const media = new Map<string, Uint8Array>();
  const keep = (assets: Array<{ assetId: string; bytes: ArrayBuffer | Uint8Array }> = []) => {
    for (const asset of assets) media.set(asset.assetId, new Uint8Array(asset.bytes));
  };
  keep(client.state.frame?.media);
  client.subscribeMedia(keep);
  return {
    pages: slides.length,
    async capture(index: number) {
      // Activation waits for the slide's default faces, as editor navigation does; a thumbnail would not.
      const slideId = slides[index].id;
      await client.dispatch({ kind: 'activateSlide', payload: { slideId } });
      const frame = client.state.frame;
      if (frame?.slideId !== slideId) throw new Error(`slide ${index + 1} did not activate`);
      keep(frame.media);
      const list = decodeCommandDisplayList(frame);
      const canvas = document.createElement('canvas');
      sizeCanvasForSlide(canvas, list, 150 / 96, 1);
      const resolver = images.resolve({ media });
      try {
        await paintSlide(canvas.getContext('2d')!, list, 150 / 96, 1, { resolveImage: resolver });
        return canvas.toDataURL('image/png');
      } finally {
        resolver.release();
      }
    },
  };
}
