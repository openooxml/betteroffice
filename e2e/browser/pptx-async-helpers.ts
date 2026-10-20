import { expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

export async function inspectDeck(page: Page, bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const files = Object.values(zip.files).filter((file) => !file.dir && /\.(xml|rels)$/.test(file.name));
  const parts = Object.fromEntries(await Promise.all(files.map(async (file) => [file.name, await file.async('string')])));
  return page.evaluate((parts) => {
    const parse = (name: string) => {
      const document = new DOMParser().parseFromString(parts[name]!, 'application/xml');
      if (document.querySelector('parsererror')) throw new Error(`Invalid XML: ${name}`);
      return document;
    };
    const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS('*', name));
    const relationships = all(parse('ppt/_rels/presentation.xml.rels'), 'Relationship');
    return all(parse('ppt/presentation.xml'), 'sldId').map((slide) => {
      const relationship = relationships.find((entry) => entry.getAttribute('Id') === slide.getAttributeNS(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'))!;
      const target = relationship.getAttribute('Target')!;
      const path = new URL(target, 'https://fixture.invalid/ppt/presentation.xml').pathname.slice(1);
      const document = parse(path);
      return {
        id: slide.getAttribute('id'),
        ids: all(document, 'cNvPr').map((shape) => shape.getAttribute('id')),
        shapes: all(document, 'sp').map((shape) => {
          const paragraphs = all(shape, 'p').map((paragraph) => all(paragraph, 't').map((text) => text.textContent ?? '').join(''));
          return {
            id: all(shape, 'cNvPr')[0]!.getAttribute('id'),
            text: paragraphs.join('\n'),
            paragraphs,
            x: Number(all(shape, 'off')[0]?.getAttribute('x')),
            y: Number(all(shape, 'off')[0]?.getAttribute('y')),
            runs: all(shape, 'r').map((run) => {
              const style = all(run, 'rPr')[0];
              return { text: all(run, 't').map((text) => text.textContent ?? '').join(''),
                bold: ['1', 'true'].includes(style?.getAttribute('b') ?? ''),
                italic: ['1', 'true'].includes(style?.getAttribute('i') ?? ''),
                underline: style?.getAttribute('u') ?? null, size: style?.getAttribute('sz') ?? null,
                color: style && all(style, 'srgbClr')[0]?.getAttribute('val')?.toUpperCase() || null };
            }),
          };
        }),
      };
    });
  }, parts);
}

export type SavedDeck = Awaited<ReturnType<typeof inspectDeck>>;

export const slideCanvas = (page: Page) => page.getByTestId('pptx-slide-canvas');

export async function painted(page: Page) {
  return slideCanvas(page).evaluate((canvas: HTMLCanvasElement) => canvas.toDataURL());
}

export async function expectRepaint(page: Page, previous: string) {
  await expect.poll(() => painted(page)).not.toBe(previous);
}
