import type { DocxEchoGlyphLine, DocxEchoGlyphSet } from '@betteroffice/docx/peerless';

const alphabet = Array.from({ length: 95 }, (_, index) => String.fromCharCode(32 + index)).join('');

/** Ten pixels per printable ASCII character; each outline starts at its own code point. */
export const testEchoGlyphSet: DocxEchoGlyphSet = {
  key: 'test', fontId: 1, alphabet, horizontal: 0.01,
  glyphIds: [...alphabet].map((char) => char.charCodeAt(0)), advances: [...alphabet].map(() => 1000),
  kerning: [], unsafePairs: [],
  outlines: Object.fromEntries([...alphabet].map((char) => [char.charCodeAt(0),
    JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: char.charCodeAt(0), y: 0 }, { t: 'L', x: 0, y: 700 }, { t: 'Z' }] })])),
};

export function testEchoGlyphLine(overrides: Partial<DocxEchoGlyphLine> = {}): DocxEchoGlyphLine {
  return { glyphSet: 'test', fontId: 1, size: 12, color: '#000', x: 0, y: 14, next: '', right: 1000,
    device: { x: 0, y: 0, width: 100, height: 20 }, pixelRatio: 1, ...overrides };
}

/** Replaces Path2D with a recorder whose `glyph` is the code point its outline starts at. */
export function installGlyphPaths(): () => void {
  const scope = globalThis as { Path2D?: unknown };
  const previous = scope.Path2D;
  scope.Path2D = class { glyph?: number; moveTo(x: number) { this.glyph ??= x; } lineTo() {} closePath() {} };
  return () => { scope.Path2D = previous; };
}

export function glyphChar(path: unknown): string {
  return String.fromCodePoint((path as { glyph: number }).glyph);
}
