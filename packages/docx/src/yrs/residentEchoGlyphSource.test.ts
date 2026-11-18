import { expect, test } from 'bun:test';
import type { DocxCommandEchoContext } from './residentCommandProtocol';
import { type EchoGlyphEngine, EchoGlyphSources } from './residentEchoGlyphSource';
import { DOCX_ECHO_ALPHABET } from './residentEchoGlyphs';

test('retained glyph providers read shipped sets and cached ink without building', () => {
  let shapes = 0;
  let outlines = 0;
  const sources = new EchoGlyphSources({
    echoGlyphSetJson() {
      shapes += 1;
      return JSON.stringify({ horizontal: 0.01, glyphIds: [1], advances: [1000], kerning: [], unsafePairs: [] });
    },
    outlineGlyphJson() {
      outlines += 1;
      return JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 500, y: 700 }, { t: 'Z' }] });
    },
  }, 1);
  const retained = sources.provider(1, true);
  expect(retained.glyphSet(1, 12, false)).toBeUndefined();
  expect(retained.glyphInk(1, 1)).toBeUndefined();
  expect([shapes, outlines]).toEqual([0, 0]);
  const built = sources.provider(1).glyphSet(1, 12, false)!;
  expect(built).toBeDefined();
  expect(retained.glyphSet(1, 12, false)).toBeUndefined();
  sources.sent([built.glyphs.set]);
  expect(retained.glyphSet(1, 12, false)).toBe(built);
  const counts = [shapes, outlines];
  expect(retained.glyphInk(1, 1)).toBeDefined();
  expect(retained.glyphSet(1, 13, false)).toBeUndefined();
  expect(retained.glyphInk(1, 2)).toBeUndefined();
  expect([shapes, outlines]).toEqual(counts);
});

/** Glyph ids are code points; each character the call adds kerns against the first one. */
function growingEngine() {
  const calls: { alphabet: string; known?: number }[] = [];
  const renumbered = new Set<string>();
  const engine: EchoGlyphEngine = {
    echoGlyphSetJson(_font: number, _size: number, _kern: boolean, alphabet: string, _language?: string, known?: number) {
      calls.push({ alphabet, known });
      const chars = [...alphabet];
      return JSON.stringify({
        horizontal: 0.01,
        glyphIds: chars.map((char) => renumbered.has(char) ? 9999 : char.codePointAt(0)!),
        advances: chars.map(() => 500),
        kerning: chars.slice(known ?? 0).map((_, index) => [(known ?? 0) + index, 0, -1]),
        unsafePairs: [],
      });
    },
    outlineGlyphJson() {
      return JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 400, y: 700 }, { t: 'Z' }] });
    },
  };
  return { engine, calls, renumbered };
}

const naming = (key: string) => ({ glyphs: { faces: [{ glyphSet: key }] } }) as unknown as DocxCommandEchoContext;

test('a set grows by what its new characters add, sent as growth only to a main thread holding its base', () => {
  const { engine, calls } = growingEngine();
  const sources = new EchoGlyphSources(engine, 1);
  const provider = sources.provider(1);
  const base = provider.glyphSet(1, 12, false, 'ko-KR')!;
  const known = [...DOCX_ECHO_ALPHABET].length;
  expect(calls).toEqual([{ alphabet: DOCX_ECHO_ALPHABET, known: undefined }]);
  expect(sources.unsent(naming(base.key))).toEqual([base.glyphs.set]);
  sources.sent([base.glyphs.set]);
  // characters it holds cost nothing
  expect(provider.glyphSet(1, 12, false, 'ko-KR', 'abc')).toBe(base);
  expect(calls.length).toBe(1);

  const grown = provider.glyphSet(1, 12, false, 'ko-KR', '가나가')!;
  expect(calls[1]).toEqual({ alphabet: `${DOCX_ECHO_ALPHABET}가나`, known });
  const [growth] = sources.unsent(naming(grown.key));
  expect(growth!.extends).toBe(base.key);
  expect(Object.keys(growth!.outlines).sort()).toEqual(['가', '나'].map((char) => String(char.codePointAt(0))).sort());
  expect(growth!.kerning).toEqual([[known, 0, -1], [known + 1, 0, -1]]);
  expect(growth!.glyphIds.length).toBe(known + 2);
  // the whole set stands alone for a receiver without the base
  const whole = grown.glyphs.set;
  expect(whole.extends).toBeUndefined();
  expect(whole.kerning.length).toBe(known + 2);
  expect(Object.keys(whole.outlines).length).toBe(known + 2);
  expect(grown.glyphs.chars.has('나')).toBe(true);
  // a context names the newest set; the replaced one is never sent again
  expect(sources.unsent(naming(base.key))).toEqual([]);

  const retained = sources.provider(1, true);
  expect(retained.glyphSet(1, 12, false, 'ko-KR', '가')).toBeUndefined();
  sources.sent([growth!]);
  expect(retained.glyphSet(1, 12, false, 'ko-KR', '가')).toBe(grown);
  expect(retained.glyphSet(1, 12, false, 'ko-KR', '다')).toBeUndefined();
  // a main thread that may have missed frames gets every set whole
  sources.resend();
  expect(sources.unsent(naming(grown.key))).toEqual([whole]);
  expect(calls.length).toBe(2);
});

test('a set whose earlier glyphs shape differently, or that would grow too large, is built again whole', () => {
  const { engine, calls, renumbered } = growingEngine();
  const sources = new EchoGlyphSources(engine, 1);
  const provider = sources.provider(1);
  sources.sent([provider.glyphSet(1, 12, false, undefined)!.glyphs.set]);
  renumbered.add('a');
  const rebuilt = provider.glyphSet(1, 12, false, undefined, '가')!;
  expect(calls.map((call) => call.known)).toEqual([undefined, [...DOCX_ECHO_ALPHABET].length, undefined]);
  const [sent] = sources.unsent(naming(rebuilt.key));
  expect(sent!.extends).toBeUndefined();
  expect(sent!.glyphIds[DOCX_ECHO_ALPHABET.indexOf('a')]).toBe(9999);
  sources.sent([sent!]);

  const many = Array.from({ length: 1000 }, (_, index) => String.fromCharCode(0xac00 + index)).join('');
  const large = provider.glyphSet(1, 12, false, undefined, many)!;
  expect(calls.at(-1)).toEqual({ alphabet: DOCX_ECHO_ALPHABET + many, known: undefined });
  // the rebuilt set starts again from the base alphabet and the line's own characters
  expect(large.glyphs.chars.has('가')).toBe(true);
  expect(large.glyphs.set.alphabet.length).toBe(DOCX_ECHO_ALPHABET.length + 1000);
});
