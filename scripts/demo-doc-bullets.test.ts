import { describe, expect, test } from 'bun:test';
import { execFileSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import JSZip from 'jszip';

const ROOT = path.resolve(import.meta.dir, '..');
const SYMBOL_FONT = /^(symbol|wingdings( [123])?)$/i;

const decode = (value: string) =>
  value
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');

const attribute = (attrs: string, name: string) => {
  const match = new RegExp(`\\b${name}="([^"]*)"`).exec(attrs);
  return match ? decode(match[1]) : undefined;
};

async function generatedNumbering(): Promise<string> {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-doc-bullets-'));
  const built = path.join(directory, 'betteroffice-demo.docx');
  try {
    execFileSync('bun', [path.join(ROOT, 'scripts/create-demo-doc.ts'), built], { cwd: ROOT, stdio: 'pipe' });
    const zip = await JSZip.loadAsync(fs.readFileSync(built));
    return (await zip.file('word/numbering.xml')?.async('string')) ?? '';
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
}

describe('the demo document numbering', () => {
  test('every symbol-font level text uses the symbol private-use codes', async () => {
    const numbering = await generatedNumbering();
    const levels = [...numbering.matchAll(/<w:lvl\b([^>]*)>([\s\S]*?)<\/w:lvl>/g)];
    expect(levels.length).toBeGreaterThan(0);
    const symbolLevels: { level: string; text: string; font: string }[] = [];
    for (const [, lvlAttrs, body] of levels) {
      const text = attribute(/<w:lvlText\b([^>]*)\/>/.exec(body)?.[1] ?? '', 'w:val');
      const fonts = /<w:rFonts\b([^>]*)\/>/.exec(/<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(body)?.[1] ?? '')?.[1] ?? '';
      const font = attribute(fonts, 'w:ascii') ?? attribute(fonts, 'w:hAnsi');
      if (text === undefined || !font || !SYMBOL_FONT.test(font.trim())) continue;
      symbolLevels.push({ level: attribute(lvlAttrs, 'w:ilvl') ?? '?', text, font });
    }
    expect(symbolLevels.length).toBeGreaterThan(0);
    const outside = symbolLevels.flatMap(({ level, text, font }) => {
      const codes = [...text.replace(/%\d/g, '')].map((char) => char.codePointAt(0)!);
      if (!codes.length) return [`level ${level} ${font}: empty`];
      return codes
        .filter((code) => code < 0xf020 || code > 0xf0ff)
        .map((code) => `level ${level} ${font}: U+${code.toString(16).toUpperCase().padStart(4, '0')}`);
    });
    expect(outside).toEqual([]);
  });

  test("the first bullet level is Word's Symbol bullet", async () => {
    const numbering = await generatedNumbering();
    const first = /<w:abstractNum w:abstractNumId="0">[\s\S]*?<w:lvl w:ilvl="0">([\s\S]*?)<\/w:lvl>/.exec(numbering)?.[1] ?? '';
    expect(attribute(/<w:lvlText\b([^>]*)\/>/.exec(first)?.[1] ?? '', 'w:val')).toBe('');
    expect(first).toContain('<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>');
  });
});
