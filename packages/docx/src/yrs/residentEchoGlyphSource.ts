import type { DocxCommandEchoContext, DocxEchoGlyphSet } from './residentCommandProtocol';
import type { EchoGlyphProvider, EchoGlyphSource } from './residentCommandViewProjection';
import { DOCX_ECHO_ALPHABET, indexEchoGlyphSet, outlineBounds } from './residentEchoGlyphs';

/** @internal */
export interface EchoGlyphEngine {
  echoGlyphSetJson(fontId: number, size: number, kern: boolean, alphabet: string): string;
  /** `syntheticBold` asks for the widened outlines a regular face is painted with instead. */
  outlineGlyphJson(fontId: number, glyphId: number, syntheticBold?: boolean): string;
}

type InkBox = { upem: number; xMin: number; xMax: number; yMin: number; yMax: number };

/** Shapes advance the same for both weights, so only the outlines and their ink differ. */
function setKey(generation: number, fontId: number, size: number, kern: boolean, syntheticBold: boolean): string {
  return `${generation}:${fontId}:${size}:${kern ? 1 : 0}${syntheticBold ? 'synthetic' : ''}`;
}

function inkKey(fontId: number, glyphId: number, syntheticBold: boolean): string {
  return `${fontId}:${glyphId}${syntheticBold ? ':synthetic' : ''}`;
}

/** @internal The echo glyph sets of one font store, and which of them the main thread was sent. */
export class EchoGlyphSources {
  private readonly sets = new Map<string, { source: EchoGlyphSource; payload: DocxEchoGlyphSet } | null>();
  private readonly inks = new Map<string, InkBox | null | undefined>();
  private readonly shipped = new Set<string>();

  constructor(private readonly engine: EchoGlyphEngine, private readonly generation: number) {}

  /** Sends every set again, as after a frame the main thread may not have presented. */
  resend(): void { this.shipped.clear(); }

  provider(pixelRatio: number, retainedOnly = false): EchoGlyphProvider {
    return {
      pixelRatio,
      glyphSet: (fontId, size, kern, syntheticBold = false) => {
        const key = setKey(this.generation, fontId, size, kern, syntheticBold);
        if (!retainedOnly) return this.glyphSet(key, fontId, size, kern, syntheticBold);
        const entry = this.sets.get(key);
        return entry && this.shipped.has(entry.payload.key) ? entry.source : undefined;
      },
      glyphInk: (fontId, glyphId, syntheticBold = false) => retainedOnly
        ? this.inks.get(inkKey(fontId, glyphId, syntheticBold))
        : this.glyphInk(fontId, glyphId, syntheticBold),
    };
  }

  unsent(context: DocxCommandEchoContext | undefined): DocxEchoGlyphSet[] {
    const key = context?.glyphs?.glyphSet;
    const entry = key === undefined ? undefined : this.sets.get(key);
    return entry && !this.shipped.has(entry.payload.key) ? [entry.payload] : [];
  }

  sent(sets: readonly DocxEchoGlyphSet[]): void {
    for (const set of sets) this.shipped.add(set.key);
  }

  private glyphSet(key: string, fontId: number, size: number, kern: boolean, syntheticBold: boolean): EchoGlyphSource | undefined {
    if (!this.sets.has(key)) this.sets.set(key, this.build(key, fontId, size, kern, syntheticBold));
    return this.sets.get(key)?.source;
  }

  private build(key: string, fontId: number, size: number, kern: boolean, syntheticBold: boolean) {
    try {
      const shaped = JSON.parse(this.engine.echoGlyphSetJson(fontId, size, kern, DOCX_ECHO_ALPHABET)) as Omit<
        DocxEchoGlyphSet, 'key' | 'fontId' | 'alphabet' | 'outlines'>;
      const outlines: Record<string, string> = {};
      const ink = { upem: 0, xMin: 0, yMin: 0, yMax: 0, overhang: 0 };
      shaped.glyphIds.forEach((id, index) => {
        if (!id) return;
        // widened outlines carry the synthetic weight, so they also set the repaint band
        const json = outlines[id] ??= this.engine.outlineGlyphJson(fontId, id, syntheticBold);
        const box = this.glyphInk(fontId, id, syntheticBold, json);
        if (box === undefined) throw new Error('Echo glyph outline is unavailable');
        // a blank glyph has no ink box, so the units per em are read from the outline itself
        const outline = JSON.parse(json) as { upem: number };
        ink.upem = outline.upem;
        if (!box) return;
        ink.xMin = Math.min(ink.xMin, box.xMin);
        ink.yMin = Math.min(ink.yMin, box.yMin);
        ink.yMax = Math.max(ink.yMax, box.yMax);
        ink.overhang = Math.max(ink.overhang, box.xMax - shaped.advances[index]!);
      });
      if (!ink.upem) return null;
      const payload: DocxEchoGlyphSet = { key, fontId, ...(syntheticBold ? { syntheticBold } : {}),
        alphabet: DOCX_ECHO_ALPHABET, ...shaped, outlines };
      return { source: { key, glyphs: indexEchoGlyphSet(payload), ink }, payload };
    } catch {
      return null;
    }
  }

  /** Null for a glyph without ink, undefined when its outline cannot be read. */
  private glyphInk(fontId: number, glyphId: number, syntheticBold: boolean, outline?: string): InkBox | null | undefined {
    const key = inkKey(fontId, glyphId, syntheticBold);
    if (!this.inks.has(key)) {
      let box: InkBox | null | undefined;
      try {
        const json = outline ?? this.engine.outlineGlyphJson(fontId, glyphId, syntheticBold);
        box = outlineBounds(json) ?? null;
      } catch {
        box = undefined;
      }
      this.inks.set(key, box);
    }
    return this.inks.get(key);
  }
}
