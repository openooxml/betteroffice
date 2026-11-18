import type { DocxCommandEchoContext, DocxEchoGlyphSet } from './residentCommandProtocol';
import type { EchoGlyphProvider, EchoGlyphSource } from './residentCommandViewProjection';
import { DOCX_ECHO_ALPHABET, indexEchoGlyphSet, outlineBounds } from './residentEchoGlyphs';

/** @internal */
export interface EchoGlyphEngine {
  echoGlyphSetJson(fontId: number, size: number, kern: boolean, alphabet: string, language?: string, known?: number): string;
  outlineGlyphJson(fontId: number, glyphId: number): string;
}

type InkBox = { upem: number; xMin: number; xMax: number; yMin: number; yMax: number };
type Shaped = Omit<DocxEchoGlyphSet, 'key' | 'extends' | 'fontId' | 'alphabet' | 'outlines'>;

/** A set grows past this many characters only by starting over from what its line needs. */
const MAX_ALPHABET = 1024;

interface SetEntry {
  key: string;
  /** Code points of the alphabet, and the same as a set. */
  alphabet: string[];
  has: Set<string>;
  source: EchoGlyphSource;
  /** Everything the main thread needs when it holds no earlier set of this font. */
  payload: DocxEchoGlyphSet;
  /** What the main thread needs when it holds the set this one grew from (`growth.extends`). */
  growth?: DocxEchoGlyphSet;
}

/** @internal The echo glyph sets of one font store, and which of them the main thread was sent. */
export class EchoGlyphSources {
  /** The newest set of each font, size, kerning gate and language; null where it cannot echo. */
  private readonly sets = new Map<string, SetEntry | null>();
  private readonly byKey = new Map<string, SetEntry>();
  private readonly inks = new Map<string, InkBox | null | undefined>();
  private readonly shipped = new Set<string>();
  private version = 0;

  constructor(private readonly engine: EchoGlyphEngine, private readonly generation: number) {}

  /** Sends every set again, as after a frame the main thread may not have presented. */
  resend(): void { this.shipped.clear(); }

  provider(pixelRatio: number, retainedOnly = false): EchoGlyphProvider {
    return {
      pixelRatio,
      glyphSet: (fontId, size, kern, language, chars = '') => {
        if (!retainedOnly) return this.glyphSet(fontId, size, kern, language, chars);
        const entry = this.sets.get(this.base(fontId, size, kern, language));
        return entry && this.shipped.has(entry.key) && [...chars].every((char) => entry.has.has(char))
          ? entry.source : undefined;
      },
      glyphInk: (fontId, glyphId) => retainedOnly ? this.inks.get(`${fontId}:${glyphId}`) : this.glyphInk(fontId, glyphId),
    };
  }

  /** The sets `context` paints with that the main thread was not sent, each as little as it needs. */
  unsent(context: DocxCommandEchoContext | undefined): DocxEchoGlyphSet[] {
    const line = context?.glyphs;
    const keys = line ? line.faces?.map((face) => face.glyphSet) ?? [line.glyphSet] : [];
    const sets: DocxEchoGlyphSet[] = [];
    for (const key of new Set(keys)) {
      const entry = this.byKey.get(key);
      if (!entry || this.shipped.has(key)) continue;
      const base = entry.growth?.extends;
      sets.push(entry.growth && base !== undefined && this.shipped.has(base) ? entry.growth : entry.payload);
    }
    return sets;
  }

  sent(sets: readonly DocxEchoGlyphSet[]): void {
    for (const set of sets) this.shipped.add(set.key);
  }

  private base(fontId: number, size: number, kern: boolean, language: string | undefined): string {
    return `${fontId}:${size}:${kern ? 1 : 0}:${language ?? ''}`;
  }

  private glyphSet(fontId: number, size: number, kern: boolean, language: string | undefined, chars: string): EchoGlyphSource | undefined {
    const base = this.base(fontId, size, kern, language);
    const current = this.sets.get(base);
    if (current === null) return undefined;
    const missing = [...new Set(chars)].filter((char) => !current?.has.has(char));
    if (current && !missing.length) return current.source;
    const grown = current && current.alphabet.length + missing.length <= MAX_ALPHABET
      ? this.build(fontId, size, kern, language, [...current.alphabet, ...missing], current)
      : this.build(fontId, size, kern, language, [...new Set([...DOCX_ECHO_ALPHABET, ...chars])]);
    // a set that cannot grow keeps echoing what it has
    if (grown === null && current) return current.source;
    this.sets.set(base, grown);
    if (grown) this.byKey.set(grown.key, grown);
    // contexts name the newest set; the main thread drops a set once the one growing it arrives
    if (current) this.byKey.delete(current.key);
    return grown?.source;
  }

  private build(
    fontId: number, size: number, kern: boolean, language: string | undefined, alphabet: string[], earlier?: SetEntry
  ): SetEntry | null {
    try {
      const known = earlier?.alphabet.length ?? 0;
      const shaped = JSON.parse(this.engine.echoGlyphSetJson(fontId, size, kern, alphabet.join(''), language,
        known || undefined)) as Shaped;
      if (earlier) {
        const before = earlier.payload;
        // pairs only grow from what the earlier set paired; a glyph it could not echo stays out
        if (before.glyphIds.some((id, index) => shaped.glyphIds[index] !== id && (id === 0 || shaped.glyphIds[index] !== 0)) ||
          before.horizontal !== shaped.horizontal) return this.build(fontId, size, kern, language, alphabet);
      }
      const outlines: Record<string, string> = { ...earlier?.payload.outlines };
      const added: Record<string, string> = {};
      const ink = { upem: 0, xMin: 0, yMin: 0, yMax: 0, overhang: 0 };
      shaped.glyphIds.forEach((id, index) => {
        if (!id) return;
        let json = outlines[id];
        if (json === undefined) json = outlines[id] = added[id] = this.engine.outlineGlyphJson(fontId, id);
        const box = this.glyphInk(fontId, id, json);
        if (box === undefined) throw new Error('Echo glyph outline is unavailable');
        ink.upem = (JSON.parse(json) as { upem: number }).upem;
        if (!box) return;
        ink.xMin = Math.min(ink.xMin, box.xMin);
        ink.yMin = Math.min(ink.yMin, box.yMin);
        ink.yMax = Math.max(ink.yMax, box.yMax);
        ink.overhang = Math.max(ink.overhang, box.xMax - shaped.advances[index]!);
      });
      if (!ink.upem) return null;
      const key = `${this.generation}:${fontId}:${size}:${kern ? 1 : 0}:${language ?? ''}:${++this.version}`;
      const text = alphabet.join('');
      const pairs = earlier ? {
        kerning: [...earlier.payload.kerning, ...shaped.kerning],
        unsafePairs: [...earlier.payload.unsafePairs, ...shaped.unsafePairs],
      } : {};
      // the whole set names no base: a receiver without the earlier one takes it as it is
      const payload: DocxEchoGlyphSet = { key, fontId, alphabet: text, ...shaped, ...pairs, outlines };
      const growth: DocxEchoGlyphSet | undefined = earlier
        ? { key, extends: earlier.key, fontId, alphabet: text, ...shaped, outlines: added } : undefined;
      return { key, alphabet, has: new Set(alphabet), payload, ...(growth ? { growth } : {}),
        source: { key, glyphs: indexEchoGlyphSet(payload), ink } };
    } catch {
      return null;
    }
  }

  /** Null for a glyph without ink, undefined when its outline cannot be read. */
  private glyphInk(fontId: number, glyphId: number, outline?: string): InkBox | null | undefined {
    const key = `${fontId}:${glyphId}`;
    if (!this.inks.has(key)) {
      let box: InkBox | null | undefined;
      try {
        const json = outline ?? this.engine.outlineGlyphJson(fontId, glyphId);
        box = outlineBounds(json) ?? null;
      } catch {
        box = undefined;
      }
      this.inks.set(key, box);
    }
    return this.inks.get(key);
  }
}
