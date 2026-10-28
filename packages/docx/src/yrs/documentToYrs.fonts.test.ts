import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import type { LayoutBlock } from '../layout/pagination/types';
import type { DisplayList, GlyphRunPrimitive } from '../layout/render/displayList';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const FONTS = '<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>';

function fontPackage(latin: string | undefined, fonts: string, prefix = '', body?: string): Uint8Array<ArrayBuffer> {
  const parts = new Map([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      (latin !== undefined ? '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' : '') + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      (latin !== undefined ? `<Relationship Id="theme" Type="${R}/theme" Target="theme/theme1.xml"/>` : '') + '</Relationships>'],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body ?? `<w:p>${prefix}<w:r><w:rPr>${fonts}</w:rPr><w:t>Title</w:t></w:r></w:p>`}<w:sectPr/></w:body></w:document>`],
  ]);
  if (latin !== undefined) parts.set('word/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="${latin}"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`);
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([...parts].map(([name, xml]) => [name, toBytes(xml)]))));
}

function textBox(kind: 'drawingml' | 'vml', fonts: string): string {
  const content = `<w:p><w:r><w:rPr>${fonts}</w:rPr><w:t>Box</w:t></w:r></w:p>`;
  if (kind === 'vml') {
    return `<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"><v:shape id="Text Box 1" o:spid="_x0000_s1026" style="position:absolute;margin-left:36pt;margin-top:72pt;width:144pt;height:72pt;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;v-text-anchor:top" filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent>${content}</w:txbxContent></v:textbox><w10:wrap type="none"/></v:shape></w:pict></w:r>`;
  }
  return `<w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="41" name="Text Box 41"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

beforeAll(() => preloadEditWasm(new Uint8Array(
  readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))
)));

function assertThemeFontPaint(session: YrsSession, document: Document, texts: string[]): void {
  const themeFont = session.registerFont(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../../../fonts/assets/Carlito-Regular.ttf')
  )));
  const explicitFont = session.registerFont(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../../../fonts/assets/Caladea-Regular.ttf')
  )));
  expect(themeFont).not.toBe(explicitFont);
  const request = buildResidentRegionLayoutRequest(document, 24, {});
  request.measurement = {
    fontChains: { 'calibri|0|0': [themeFont], 'cambria|0|0': [explicitFont] },
    defaults: { fontSize: 11, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
  const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request))) as Array<{ family: string }>;
  expect(requirements.some(font => font.family === 'Calibri')).toBe(true);
  expect(requirements.some(font => font.family === 'Cambria')).toBe(false);
  const layout = session.layoutDocumentWithRegionsJson(JSON.stringify(request));
  const display = JSON.parse(session.buildDisplayListJson(layout)) as DisplayList;
  const glyphs = display.pages.flatMap(page => page.primitives).filter(
    (primitive): primitive is GlyphRunPrimitive => primitive.kind === 'glyphRun' && texts.includes(primitive.text)
  );
  for (const text of texts) expect(glyphs.some(run => run.text === text)).toBe(true);
  expect(glyphs.every(run => run.fontId === themeFont)).toBe(true);
}

for (const path of ['native', 'typescript'] as const) {
  for (const [name, latin, fonts, effective] of [
    ['paired', 'Calibri', FONTS, 'Calibri'],
    ['theme-only', 'Calibri', '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>', 'Calibri'],
    ['explicit-only', 'Calibri', '<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/>', 'Cambria'],
    ['default major', undefined, FONTS, 'Calibri Light'],
    ['default minor', undefined, FONTS.replaceAll('majorHAnsi', 'minorHAnsi'), 'Calibri'],
    ['empty theme slot', '', FONTS, 'Cambria'],
    ['unknown reference', 'Calibri', FONTS.replaceAll('majorHAnsi', 'unknown'), 'Cambria'],
  ] as const) {
    test(`${path} ${name} fonts resolve only in layout`, async () => {
      const bytes = fontPackage(latin, fonts);
      const document = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (path === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, document);
        const original = session.encodeState();
        const segments = session.storySegments('body');
        const authored = name === 'theme-only' ? 'Calibri' : 'Cambria';
        expect(segments.find(segment => segment.kind === 'text')?.attributes.fontFamily).toMatchObject({
          ascii: authored, hAnsi: authored,
        });
        const request = buildResidentRegionLayoutRequest(document, 24, {});
        const blocks = session.yrsBlocksForStory('body', request.renderEnv) as LayoutBlock[];
        const runs = blocks.flatMap(block => block.kind === 'paragraph' ? block.runs : []);
        expect(runs).toHaveLength(1);
        expect(runs[0]).toMatchObject({
          kind: 'text', text: 'Title', fontFamily: effective,
          fontSlots: { ascii: effective, hAnsi: effective },
        });
        if (name === 'paired') assertThemeFontPaint(session, document, ['Title']);
        expect(session.encodeState()).toEqual(original);
      } finally {
        session.destroy();
      }
    });
  }

  for (const kind of ['drawingml', 'vml'] as const) {
    for (const [authored, fonts] of [
      ['paired', FONTS],
      ['theme-only', '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>'],
    ] as const) {
      test(`${path} ${kind} ${authored} text box fonts resolve only in layout`, async () => {
        const bytes = fontPackage('Calibri', fonts, textBox(kind, fonts));
        const document = await parseDocx(bytes.buffer, { preloadFonts: false });
        const session = await createYrsSession();
        try {
          if (path === 'native') session.seedFromDocx(bytes);
          else documentToYrs(session, document);
          const original = session.encodeState();
          const request = buildResidentRegionLayoutRequest(document, 24, {});
          const blocks = session.yrsBlocksForStory('body', request.renderEnv) as LayoutBlock[];
          const runs = blocks.flatMap(block => block.kind === 'shape'
            ? (block.innerText ?? []).flatMap(paragraph => paragraph.runs)
            : block.kind === 'textBox' ? block.content.flatMap(paragraph => paragraph.runs) : []);
          const labels = runs.filter(run => run.kind === 'text' && run.text === 'Box');
          expect(labels).toHaveLength(1);
          expect(labels[0]).toMatchObject({
            fontFamily: 'Calibri', fontSlots: { ascii: 'Calibri', hAnsi: 'Calibri' },
          });
          if (authored === 'paired') assertThemeFontPaint(session, document, ['Box', 'Title']);
          expect(session.encodeState()).toEqual(original);
        } finally {
          session.destroy();
        }
      });
    }
  }

  for (const action of ['split', 'merge', 'field edit'] as const) {
    test(`${path} ${action} saves the authored font pair`, async () => {
      const run = `<w:r><w:rPr>${FONTS}</w:rPr><w:t>Title</w:t></w:r>`;
      const body = action === 'merge'
        ? `<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p>${run}</w:p>`
        : action === 'field edit'
          ? `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc1">${run}</w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
          : `<w:p><w:r><w:t>A</w:t></w:r>${run}</w:p>`;
      const bytes = fontPackage('Calibri', FONTS, '', body);
      const document = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession();
      const hydrated = await createYrsSession();
      try {
        if (path === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, document);
        hydrated.openDocx(bytes, false);
        hydrated.loadState(session.encodeState());
        const target = hydrated.paragraphs('body')[0]!;
        if (action === 'split') hydrated.splitParagraph({ story: 'body', paraId: target.paraId, offset: 1 });
        else if (action === 'merge') hydrated.mergeParagraphs('body', target.paraId);
        else hydrated.insertText({ story: 'body', paraId: target.paraId, offset: 1 }, '!');
        const saved = unzipContainer((await saveYrsDocx(hydrated)).bytes);
        const xml = new TextDecoder().decode(saved['word/document.xml']);
        expect([...xml.matchAll(/<w:rFonts\b[^>]*\/>/g)].map(match => match[0])).toEqual([FONTS]);
        expect([...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]).join(''))
          .toBe(action === 'field edit' ? 'T!itle' : 'ATitle');
        expect([...xml.matchAll(/<w:p\b[^>]*>/g)]).toHaveLength(action === 'split' ? 2 : 1);
        if (action === 'field edit') {
          expect(xml).toContain('TOC');
          expect(xml).toContain('_Toc1');
        }
      } finally {
        session.destroy();
        hydrated.destroy();
      }
    });
  }
}
