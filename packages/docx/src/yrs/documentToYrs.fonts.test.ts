import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import type { LayoutBlock, Run as LayoutRun, ShapeBlock } from '../layout/pagination/types';
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

function mixedSlotsPackage(fonts: string): Uint8Array<ArrayBuffer> {
  const body = `<w:p>${textBox('drawingml', fonts)}</w:p><w:p>${textBox('vml', fonts)}</w:p>` +
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    `<w:r><w:rPr>${fonts}</w:rPr><w:t>Title</w:t></w:r></w:p>`;
  const parts = new Map(Object.entries(unzipContainer(fontPackage('Calibri', fonts, '', body))));
  const types = new TextDecoder().decode(parts.get('[Content_Types].xml'));
  parts.set('[Content_Types].xml', toBytes(types.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')));
  const rels = new TextDecoder().decode(parts.get('word/_rels/document.xml.rels'));
  parts.set('word/_rels/document.xml.rels', toBytes(rels.replace('</Relationships>', `<Relationship Id="numbering" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`)));
  parts.set('word/numbering.xml', toBytes(`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr>${fonts}</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`));
  parts.set('word/theme/theme1.xml', toBytes('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

type FontSlots = 'latin' | 'eastAsia/cs' | 'all slots';

function pairedFonts(slots: FontSlots, site: string): string {
  const latin = slots !== 'eastAsia/cs';
  const other = slots !== 'latin';
  const attributes = [
    ...(latin ? [['ascii', `${site} Cambria`], ['hAnsi', `${site} Cambria`]] : []),
    ...(other ? [['eastAsia', `${site} MS Mincho`], ['cs', `${site} Arial`]] : []),
    ...(latin ? [['asciiTheme', 'majorHAnsi'], ['hAnsiTheme', 'majorHAnsi']] : []),
    ...(other ? [['eastAsiaTheme', 'minorEastAsia'], ['cstheme', 'minorBidi']] : []),
    ['hint', 'eastAsia'],
  ];
  return `<w:rFonts ${attributes.map(([name, value]) => `w:${name}="${value}"`).join(' ')}/>`;
}

function pairedFontsPackage(slots: FontSlots, theme: 'custom' | 'default'): Uint8Array<ArrayBuffer> {
  const fonts = (site: string) => pairedFonts(slots, site);
  const run = (site: string, text: string) => `<w:r><w:rPr>${fonts(site)}</w:rPr><w:t>${text}</w:t></w:r>`;
  const relationships = [
    ['styles', 'styles.xml', 'wordprocessingml.styles'],
    ['numbering', 'numbering.xml', 'wordprocessingml.numbering'],
    ['header', 'header1.xml', 'wordprocessingml.header'],
    ['footer', 'footer1.xml', 'wordprocessingml.footer'],
    ...(theme === 'custom' ? [['theme', 'theme/theme1.xml', 'theme']] : []),
  ];
  const body = `<w:p>${run('Anchor', 'A')}${run('Direct', 'Title')}</w:p>` +
    `<w:p>${run('Merge', 'Merge')}</w:p>` +
    '<w:p><w:r><w:t>Document defaults</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pStyle w:val="PairedParagraph"/></w:pPr><w:r><w:t>Paragraph style</w:t></w:r></w:p>' +
    '<w:p><w:r><w:rPr><w:rStyle w:val="PairedCharacter"/></w:rPr><w:t>Character style</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>List</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr><w:tblW w:w="2400" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p>${run('Table', 'Cell')}</w:p></w:tc></w:tr></w:tbl>` +
    `<w:p>${textBox('drawingml', fonts('DrawingML'))}</w:p>` +
    `<w:p>${textBox('vml', fonts('VML'))}</w:p>` +
    `<w:p>${run('SimplePrefix', 'BeforeSimple')}<w:fldSimple w:instr=" PAGE ">${run('SimpleField', '7')}</w:fldSimple></w:p>` +
    `<w:p>${run('ComplexPrefix', 'BeforeComplex')}<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    '<w:r><w:instrText xml:space="preserve"> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    `<w:hyperlink w:anchor="_Toc1">${run('ComplexField', 'Heading')}</w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
  const parts = new Map([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      relationships.map(([, target, type]) => `<Override PartName="/word/${target}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`).join('') + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      relationships.map(([type, target]) => `<Relationship Id="${type}" Type="${R}/${type}" Target="${target}"/>`).join('') + '</Relationships>'],
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/></w:sectPr></w:body></w:document>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${fonts('Defaults')}</w:rPr></w:rPrDefault></w:docDefaults>` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      `<w:style w:type="paragraph" w:styleId="PairedParagraph"><w:name w:val="Paired Paragraph"/><w:basedOn w:val="Normal"/><w:rPr>${fonts('ParagraphStyle')}</w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="PairedCharacter"><w:name w:val="Paired Character"/><w:rPr>${fonts('CharacterStyle')}</w:rPr></w:style></w:styles>`],
    ['word/numbering.xml', `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr>${fonts('Numbering')}</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`],
    ['word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p>${run('Header', 'Header')}</w:p></w:hdr>`],
    ['word/footer1.xml', `<w:ftr xmlns:w="${W}"><w:p>${run('Footer', 'Footer')}</w:p></w:ftr>`],
  ]);
  if (theme === 'custom') {
    parts.set('word/theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface="Yu Gothic"/><a:cs typeface="Tahoma"/></a:majorFont><a:minorFont><a:latin typeface="Carlito"/><a:ea typeface="SimSun"/><a:cs typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
  }
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([...parts].map(([name, xml]) => [name, toBytes(xml)]))));
}

function rFontsInXml(xml: string): string[] {
  return [...xml.matchAll(/<w:rFonts\b[^>]*(?:\/>|>[\s\S]*?<\/w:rFonts>)/g)].map(match => match[0]);
}

function rFontsByPart(bytes: Uint8Array): Record<string, string[]> {
  return Object.fromEntries(Object.entries(unzipContainer(bytes)).flatMap(([part, content]) => {
    if (!part.endsWith('.xml')) return [];
    const xml = new TextDecoder().decode(content);
    const fonts = rFontsInXml(xml);
    return fonts.length ? [[part, fonts]] : [];
  }));
}

function rFontsAttributes(parts: Record<string, string[]>): Record<string, Record<string, string>[]> {
  return Object.fromEntries(Object.entries(parts).map(([part, fonts]) => [part, fonts.map(font =>
    Object.fromEntries([...font.matchAll(/([\w:.-]+)="([^"]*)"/g)].map(match => [match[1], match[2]]))
  )]));
}

function textInXml(xml: string): string {
  return [...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(text => text[1]).join('');
}

function authoredRFontsAfterEdit(
  bytes: Uint8Array,
  slots: FontSlots,
  insertedPrefixes: readonly string[]
): Record<string, Record<string, string>[]> {
  const parts = rFontsByPart(bytes);
  const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  const sourceFonts = (site: string) => rFontsAttributes({ run: [pairedFonts(slots, site)] }).run[0]!;
  const authored = xml.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (run, offset: number) => {
    const fonts = rFontsInXml(run);
    if (fonts.length === 0) return run;
    if (/<w:fldChar\b/.test(run)) {
      expect(fonts).toHaveLength(1);
      expect(rFontsAttributes({ run: fonts }).run).toEqual([sourceFonts('ComplexField')]);
      return run.replace(fonts[0]!, '');
    }
    if (textInXml(run) === '!') {
      const prefix = insertedPrefixes.map(prefix => ({ prefix, offset: xml.lastIndexOf(prefix, offset) }))
        .filter(entry => entry.offset >= 0).sort((left, right) => right.offset - left.offset)[0]?.prefix;
      expect(prefix).toBeDefined();
      const expected = sourceFonts(prefix === 'BeforeSimple' ? 'SimplePrefix' : 'ComplexPrefix');
      delete expected['w:hint'];
      expect(fonts).toHaveLength(1);
      expect(rFontsAttributes({ run: fonts }).run).toEqual([expected]);
      return run.replace(fonts[0]!, '');
    }
    return run;
  });
  parts['word/document.xml'] = rFontsInXml(authored);
  return rFontsAttributes(parts);
}

function paragraphTexts(xml: string): string[] {
  return [...xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)].map(paragraph => textInXml(paragraph[1]));
}

beforeAll(() => preloadEditWasm(new Uint8Array(
  readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))
)));

function layoutRuns(blocks: readonly LayoutBlock[]): LayoutRun[] {
  return blocks.flatMap(block => {
    const runs = block.kind === 'paragraph' ? block.runs
      : block.kind === 'shape' ? (block.innerText ?? []).flatMap(paragraph => paragraph.runs)
        : block.kind === 'textBox' ? block.content.flatMap(paragraph => paragraph.runs) : [];
    return runs.flatMap(run => run.kind === 'image' && run.inlineShape
      ? layoutRuns([run.inlineShape as ShapeBlock]) : [run]);
  });
}

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
  const layout = JSON.parse(session.layoutDocumentWithRegionsJson(JSON.stringify(request)));
  const display = JSON.parse(session.buildDisplayListJson(JSON.stringify({
    ...layout, fontChains: request.measurement.fontChains,
  }))) as DisplayList;
  const painted = display.pages.flatMap(page => page.primitives).filter(
    primitive => (primitive.kind === 'glyphRun' || primitive.kind === 'text') && primitive.text.length > 0
  );
  expect(painted.map(primitive => primitive.kind)).toEqual(painted.map(() => 'glyphRun'));
  const glyphs = painted.filter((primitive): primitive is GlyphRunPrimitive => primitive.kind === 'glyphRun');
  const paintedText = glyphs.map(run => run.text).join('');
  for (const text of texts) expect(paintedText).toContain(text);
  expect(paintedText.length).toBe(texts.join('').length);
  expect(glyphs.map(run => run.fontId)).toEqual(glyphs.map(() => themeFont));
}

for (const path of ['native', 'typescript'] as const) {
  for (const [slot, fonts, ascii, hAnsi] of [
    ['hAnsi', '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Arial" w:hAnsiTheme="minorHAnsi"/>', 'Times New Roman', 'Calibri'],
    ['ascii', '<w:rFonts w:ascii="Times New Roman" w:asciiTheme="minorHAnsi" w:hAnsi="Arial"/>', 'Calibri', 'Arial'],
  ] as const) {
    test(`${path} mixed ${slot} theme fonts stay in their own run and numbering slots`, async () => {
      const bytes = mixedSlotsPackage(fonts);
      const document = await parseDocx(bytes.buffer, { preloadFonts: false });
      const source = await createYrsSession();
      const session = await createYrsSession();
      try {
        if (path === 'native') source.seedFromDocx(bytes);
        else documentToYrs(source, document);
        session.openDocx(bytes, false);
        session.loadState(source.encodeState());
        const original = session.encodeState();
        const request = buildResidentRegionLayoutRequest(document, 24, {});
        const blocks = session.yrsBlocksForStory('body', request.renderEnv) as LayoutBlock[];
        const runs = layoutRuns(blocks);
        const text = runs.filter(run => run.kind === 'text' && ['Title', 'Box'].includes(run.text));
        expect(text).toHaveLength(3);
        for (const run of text) expect(run).toMatchObject({
          fontFamily: ascii, fontSlots: { ascii, hAnsi },
        });
        const paragraph = blocks.find(block => block.kind === 'paragraph' &&
          block.runs.some(run => run.kind === 'text' && run.text === 'Title'));
        expect(paragraph).toMatchObject({ attrs: { listMarker: '1.', listMarkerFontFamily: ascii } });
        expect(session.storySegments('body').find(segment => segment.kind === 'text')?.attributes.fontFamily)
          .toMatchObject({ ascii: 'Times New Roman', hAnsi: 'Arial' });
        expect(session.encodeState()).toEqual(original);
        expect(rFontsByPart((await saveYrsDocx(session)).bytes)).toEqual(rFontsByPart(bytes));
      } finally {
        source.destroy();
        session.destroy();
      }
    });
  }

  for (const [authored, fonts, seeded] of [
    ['theme-only', '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>', 'Times New Roman'],
    ['paired', FONTS, 'Cambria'],
  ] as const) {
    for (const context of ['unknown theme', 'document theme'] as const) {
      test(`${path} source-less ${authored} fonts resolve with ${context} without changing the model`, async () => {
        const bytes = fontPackage('Times New Roman', fonts, textBox('drawingml', fonts) + textBox('vml', fonts));
        const document = await parseDocx(bytes.buffer, { preloadFonts: false });
        const source = await createYrsSession();
        const session = await createYrsSession();
        try {
          if (path === 'native') source.seedFromDocx(bytes);
          else documentToYrs(source, document);
          session.loadState(source.encodeState());
          const env = context === 'document theme'
            ? buildResidentRegionLayoutRequest(document, 24, {}).renderEnv
            : undefined;
          for (const candidate of path === 'typescript' ? [source, session] : [session]) {
            const original = candidate.encodeState();
            const blocks = candidate.yrsBlocksForStory('body', env) as LayoutBlock[];
            const runs = layoutRuns(blocks);
            const text = runs.filter(run => run.kind === 'text' && ['Title', 'Box'].includes(run.text));
            expect(text).toHaveLength(3);
            expect(text.filter(run => run.kind === 'text' && run.text === 'Title')).toHaveLength(1);
            expect(text.filter(run => run.kind === 'text' && run.text === 'Box')).toHaveLength(2);
            const effective = context === 'document theme' ? 'Times New Roman' : seeded;
            for (const run of text) expect(run).toMatchObject({
              fontFamily: effective, fontSlots: { ascii: effective, hAnsi: effective },
            });
            expect(candidate.storySegments('body').find(segment => segment.kind === 'text')?.attributes.fontFamily)
              .toMatchObject({ ascii: seeded, hAnsi: seeded });
            expect(candidate.encodeState()).toEqual(original);
          }
        } finally {
          source.destroy();
          session.destroy();
        }
      });
    }
  }

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
        if (name === 'paired' || name === 'theme-only') assertThemeFontPaint(session, document, ['Title']);
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
          const runs = layoutRuns(blocks);
          const labels = runs.filter(run => run.kind === 'text' && run.text === 'Box');
          expect(labels).toHaveLength(1);
          expect(labels[0]).toMatchObject({
            fontFamily: 'Calibri', fontSlots: { ascii: 'Calibri', hAnsi: 'Calibri' },
          });
          assertThemeFontPaint(session, document, ['Box', 'Title']);
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
        expect(rFontsInXml(xml)).toEqual(action === 'field edit' ? [FONTS, FONTS, FONTS, FONTS] : [FONTS]);
        expect([...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]).join(''))
          .toBe(action === 'field edit' ? 'T!itle' : 'ATitle');
        expect([...xml.matchAll(/<w:p\b[^>]*>/g)]).toHaveLength(action === 'split' ? 2 : 1);
        if (action === 'field edit') {
          expect(xml).toContain('TOC');
          expect(xml).toContain('_Toc1');
          const runs = [...xml.matchAll(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g)].map(match => match[0]);
          expect(runs.filter(run => /<w:fldChar\b/.test(run)).flatMap(rFontsInXml)).toEqual([FONTS, FONTS, FONTS]);
          expect(runs.filter(run => /<w:t\b/.test(run)).flatMap(rFontsInXml)).toEqual([FONTS]);
        }
      } finally {
        session.destroy();
        hydrated.destroy();
      }
    });
  }

  for (const slots of ['latin', 'eastAsia/cs', 'all slots'] as const) {
    for (const theme of ['custom', 'default'] as const) {
      for (const action of ['unchanged save', 'split merge and field-adjacent edit'] as const) {
        test(`${path} ${slots} paired fonts with ${theme} theme survive ${action} in every source location`, async () => {
          const bytes = pairedFontsPackage(slots, theme);
          const originalParts = unzipContainer(bytes);
          const original = rFontsByPart(bytes);
          expect(Object.keys(original).sort()).toEqual([
            'word/document.xml', 'word/footer1.xml', 'word/header1.xml', 'word/numbering.xml', 'word/styles.xml',
          ]);
          expect(original['word/document.xml']).toHaveLength(10);
          expect(original['word/styles.xml']).toHaveLength(3);
          expect(original['word/numbering.xml']).toHaveLength(1);
          expect(original['word/header1.xml']).toHaveLength(1);
          expect(original['word/footer1.xml']).toHaveLength(1);
          const document = await parseDocx(bytes.buffer, { preloadFonts: false });
          const source = await createYrsSession();
          const session = await createYrsSession();
          try {
            if (path === 'native') source.seedFromDocx(bytes);
            else documentToYrs(source, document);
            session.openDocx(bytes, false);
            session.loadState(source.encodeState());
            const insertedPrefixes: string[] = [];
            const assertSavedFonts = async () => {
              const saved = (await saveYrsDocx(session)).bytes;
              const savedParts = unzipContainer(saved);
              expect(savedParts['word/numbering.xml']).toEqual(originalParts['word/numbering.xml']);
              const actual = rFontsByPart(saved);
              if (action === 'unchanged save') expect(actual).toEqual(original);
              else expect(authoredRFontsAfterEdit(saved, slots, insertedPrefixes)).toEqual(rFontsAttributes(original));
              const xml = new TextDecoder().decode(savedParts['word/document.xml']);
              let expectedText = textInXml(new TextDecoder().decode(originalParts['word/document.xml']));
              for (const prefix of insertedPrefixes) expectedText = expectedText.replace(prefix, `${prefix}!`);
              expect(textInXml(xml)).toBe(expectedText);
              return xml;
            };
            if (action === 'unchanged save') {
              await assertSavedFonts();
            } else {
              const first = session.paragraphs('body').find(paragraph => paragraph.text === 'ATitle')!;
              expect(first).toBeDefined();
              const split = session.splitParagraph({ story: 'body', paraId: first.paraId, offset: 1 });
              expect(session.paragraphs('body').slice(0, 3).map(paragraph => paragraph.text)).toEqual(['A', 'Title', 'Merge']);
              expect(paragraphTexts(await assertSavedFonts()).slice(0, 3)).toEqual(['A', 'Title', 'Merge']);
              session.mergeParagraphs('body', split.secondParaId);
              expect(session.paragraphs('body').slice(0, 2).map(paragraph => paragraph.text)).toEqual(['A', 'TitleMerge']);
              expect(paragraphTexts(await assertSavedFonts()).slice(0, 2)).toEqual(['A', 'TitleMerge']);
              for (const prefix of ['BeforeSimple', 'BeforeComplex']) {
                const paragraph = session.paragraphs('body').find(entry => entry.text.startsWith(prefix))!;
                expect(paragraph).toBeDefined();
                session.insertText({ story: 'body', paraId: paragraph.paraId, offset: prefix.length }, '!');
                insertedPrefixes.push(prefix);
                const xml = await assertSavedFonts();
                expect(paragraphTexts(xml).some(text => text.startsWith(`${prefix}!`))).toBe(true);
                expect(xml).toContain('<w:fldSimple w:instr=" PAGE ">');
                expect(xml).toContain('TOC');
                expect(xml).toContain('_Toc1');
                expect([...xml.matchAll(/<w:fldChar\b[^>]*w:fldCharType="(begin|separate|end)"/g)]
                  .map(match => match[1])).toEqual(['begin', 'separate', 'end']);
              }
            }
          } finally {
            source.destroy();
            session.destroy();
          }
        });
      }
    }
  }
}
