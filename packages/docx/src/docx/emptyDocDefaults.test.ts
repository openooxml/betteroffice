import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from '../yrs';
import { documentToYrs } from '../yrs/documentToYrs';
import { parseDocx } from '.';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from './rezip/parts';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))), 30_000);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function docx(styles: string): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  parts.set('word/styles.xml', toBytes(styles));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

type Spacing = { after?: number; line?: number } | undefined;

async function spacings(bytes: Uint8Array): Promise<[Spacing, Spacing]> {
  const lowered = async (open: (session: Awaited<ReturnType<typeof createYrsSession>>) => void | Promise<void>, clientId: number) => {
    const session = await createYrsSession({ clientId });
    try {
      await open(session);
      const blocks = session.yrsBlocksForStory('body', {}) as Array<{ kind: string; attrs?: { spacing?: Spacing } }>;
      return blocks.find((block) => block.kind === 'paragraph')?.attrs?.spacing;
    } finally {
      session.destroy();
    }
  };
  const parsed = await parseDocx(bytes.slice().buffer, { preloadFonts: false });
  return [
    await lowered((session) => { session.seedFromDocx(bytes); }, 24860),
    await lowered((session) => { documentToYrs(session, parsed); }, 24861),
  ];
}

test('an empty docDefaults element leaves unstyled paragraphs without spacing in both seeders', async () => {
  const [rust, ts] = await spacings(docx(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style></w:styles>`));
  expect(ts).toEqual(rust);
  expect(rust?.after ?? 0).toBe(0);
  expect(rust?.line ?? 1).toBe(1);
});

test('a styles part without docDefaults or Normal keeps Word\'s Normal fallback in both seeders', async () => {
  const [rust, ts] = await spacings(docx(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style></w:styles>`));
  expect(ts).toEqual(rust);
  expect(rust?.after).toBeCloseTo(160 / 15, 6);
  expect(rust?.line).toBeCloseTo(259 / 240, 9);
});
