import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import type { LayoutBlock } from '../layout/pagination/types';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const FONTS = '<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>';

function fontPackage(latin: string | undefined, fonts: string): Uint8Array<ArrayBuffer> {
  const parts = new Map([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      (latin !== undefined ? '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' : '') + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      (latin !== undefined ? `<Relationship Id="theme" Type="${R}/theme" Target="theme/theme1.xml"/>` : '') + '</Relationships>'],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr>${fonts}</w:rPr><w:t>Title</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`],
  ]);
  if (latin !== undefined) parts.set('word/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="${latin}"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`);
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([...parts].map(([name, xml]) => [name, toBytes(xml)]))));
}

beforeAll(() => preloadEditWasm(new Uint8Array(
  readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))
)));

for (const path of ['native', 'typescript'] as const) {
  for (const [theme, latin, reference, effective] of [
    ['resolved', 'Calibri', 'majorHAnsi', 'Calibri'],
    ['default major', undefined, 'majorHAnsi', 'Calibri Light'],
    ['default minor', undefined, 'minorHAnsi', 'Calibri'],
    ['unresolved', '', 'majorHAnsi', 'Cambria'],
    ['unknown', 'Calibri', 'unknown', 'Cambria'],
  ] as const) {
    test(`${path} ${theme} theme uses effective fonts and saves authored pairs`, async () => {
      const authoredFonts = FONTS.replaceAll('majorHAnsi', reference);
      const bytes = fontPackage(latin, authoredFonts);
      const session = await createYrsSession();
      const hydrated = await createYrsSession();
      try {
        const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
        if (path === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, parsed);
        const paragraph = parsed.package.document.content[0];
        if (paragraph?.type !== 'paragraph' || paragraph.content[0]?.type !== 'run') throw new Error('missing run');
        expect(paragraph.content[0].formatting?.fontFamily).toEqual({
          ascii: 'Cambria', hAnsi: 'Cambria', asciiTheme: reference, hAnsiTheme: reference,
        });
        const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
        const fonts = blocks.flatMap(block => block.kind === 'paragraph'
          ? block.runs.flatMap(run => run.kind === 'text' ? [run] : []) : []);
        expect(fonts).toHaveLength(1);
        expect(fonts[0]).toMatchObject({
          fontFamily: effective,
          fontSlots: { ascii: effective, hAnsi: effective },
        });
        hydrated.openDocx(bytes, false);
        hydrated.loadState(session.encodeState());
        const target = hydrated.paragraphs('body')[0]!;
        hydrated.insertText({ story: 'body', paraId: target.paraId, offset: 5 }, '!');
        const saved = unzipContainer((await saveYrsDocx(hydrated)).bytes);
        const xml = new TextDecoder().decode(saved['word/document.xml']);
        expect(xml).toContain(authoredFonts);
        expect([...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]).join('')).toBe('Title!');
      } finally {
        session.destroy();
        hydrated.destroy();
      }
    });
  }
}
