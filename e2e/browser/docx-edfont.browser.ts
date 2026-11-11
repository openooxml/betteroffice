import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

const paragraphs = [
  { family: 'Aptos', text: 'abg Aptos draws with a stand-in face' },
  { family: 'Calibri Light', text: 'Calibri Light draws with a stand-in face' },
  { family: 'Symbol', text: 'abg' },
  { family: 'Calibri', text: 'Calibri draws with its own metric twin' },
];
const symbolDrawn = 'αβγ';

interface LineRun { kind: string; text: string; glyphs: number[] }
interface Line { y: number; runs: LineRun[] }

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    paragraphs.map(({ family, text }) => '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr>' +
      `<w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/><w:sz w:val="22"/></w:rPr>` +
      `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function firstPageLines(page: Page): Promise<Line[]> {
  return page.evaluate(() => {
    type Primitive = { kind: string; text?: string; listMarker?: boolean; baselineY?: number; glyphs?: Array<{ id: number; y: number }> };
    const mirror = document.querySelector('[data-page-index="0"] .canvas-page-mirror');
    if (!mirror) return [];
    const key = Object.keys(mirror).find(name => name.startsWith('__reactFiber$'));
    let fiber = key ? (mirror as unknown as Record<string, { return?: unknown; memoizedProps?: { page?: { primitives?: Primitive[] } } }>)[key] : undefined;
    while (fiber && !fiber.memoizedProps?.page?.primitives) fiber = fiber.return as typeof fiber;
    const lines = new Map<number, Array<{ kind: string; text: string; glyphs: number[] }>>();
    for (const primitive of fiber?.memoizedProps?.page?.primitives ?? []) {
      if ((primitive.kind !== 'glyphRun' && primitive.kind !== 'text') || primitive.listMarker || !primitive.text?.trim()) continue;
      const y = Math.round(primitive.kind === 'glyphRun' ? primitive.glyphs?.[0]?.y ?? -1 : primitive.baselineY ?? -1);
      const runs = lines.get(y) ?? [];
      runs.push({ kind: primitive.kind, text: primitive.text, glyphs: (primitive.glyphs ?? []).map(glyph => glyph.id) });
      lines.set(y, runs);
    }
    return [...lines].sort(([a], [b]) => a - b).map(([y, runs]) => ({ y, runs }));
  });
}

const compact = (text: string) => text.replace(/\s+/g, '');
const lineText = (line: Line) => compact(line.runs.map(run => run.text).join(''));

for (const worker of [true, false]) {
  test(`${worker ? 'worker' : 'in-thread'} editor draws families it does not carry with a stand-in face on one line each`, async ({ page }) => {
    const warnings: string[] = [];
    page.on('console', message => {
      if (message.type() === 'warning' && /fontRegistry/.test(message.text())) warnings.push(message.text());
    });
    await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
    await page.locator('input[type=file]').last().setInputFiles({ name: 'edfont.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await documentBytes() });
    await expect(page.locator('[data-page-index="0"] canvas').first()).toBeVisible({ timeout: 60000 });
    if (worker) await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 60000 });
    const expected = paragraphs.map(({ family, text }) => compact(family === 'Symbol' ? symbolDrawn : text));
    await expect.poll(async () => (await firstPageLines(page)).map(line => ({
      text: lineText(line),
      kinds: [...new Set(line.runs.map(run => run.kind))],
      glyphs: line.runs.every(run => run.glyphs.length > 0),
    })), { timeout: 60000, intervals: [250] }).toEqual(expected.map(text => ({ text, kinds: ['glyphRun'], glyphs: true })));
    const lines = await firstPageLines(page);
    const aptos = lines[0]!.runs.flatMap(run => run.glyphs).slice(0, 3);
    const symbol = lines[2]!.runs.flatMap(run => run.glyphs);
    expect(symbol).toHaveLength(3);
    expect(symbol).not.toEqual(aptos);
    expect(warnings.filter(text => text.includes('stands in'))).toEqual([]);
    expect(warnings.some(text => text.includes('"Calibri"'))).toBe(false);
    await expect(page.locator('.notice')).toHaveCount(0);
  });
}
