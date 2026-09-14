import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import fixture from '../../../../crates/docx-edit/tests/fixtures/toc_styles.json';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { Document, Run } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
type LoweredParagraph = { runs: Array<{
  color?: string;
  underline?: unknown;
  hyperlink?: { href: string; noDefaultStyle?: boolean };
}> };

function source(): Uint8Array<ArrayBuffer> {
  const body = fixture.paragraphStyles.map((id) =>
    `<w:p><w:pPr><w:pStyle w:val="${id}"/></w:pPr>${fixture.content}</w:p>`).join('');
  const parts = new Map<string, Uint8Array>();
  for (const [name, xml] of [
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}">${fixture.styles}</w:styles>`],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`],
  ]) parts.set(name!, toBytes(xml!));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function linkedRuns(document: Document): Run[] {
  const paragraph = document.package.document.content[0];
  if (paragraph?.type !== 'paragraph') throw new Error('missing paragraph');
  const link = paragraph.content[0];
  if (link?.type !== 'hyperlink') throw new Error('missing hyperlink');
  expect(link.anchor).toBe('heading');
  return link.children.filter((child): child is Run => child.type === 'run');
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

for (const seeder of ['projected', 'native']) {
  test(`${seeder} preserves TOC provenance through save and explicit same-color edits`, async () => {
    const bytes = source();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const env = buildResidentRegionLayoutRequest(parsed, 24, {}).renderEnv;
    expect(env.tocStyleIds).toEqual(fixture.tocStyles);
    const session = await createYrsSession({ clientId: 74502 });
    const reopened = await createYrsSession({ clientId: 74503 });
    try {
      if (seeder === 'native') {
        const opened = session.openDocx(bytes, true);
        expect(buildResidentRegionLayoutRequest(opened.document, 24, {}).renderEnv.tocStyleIds)
          .toEqual(fixture.tocStyles);
      } else documentToYrs(session, parsed);
      const request = buildResidentRegionLayoutRequest(parsed, 24, {});
      const output = JSON.parse(session.layoutDocumentWithRegionsJson(JSON.stringify({
        ...request,
        measurement: { fontChains: {}, defaults: { fontFamily: 'Calibri', fontSize: 12 } },
      }))) as { measured: Array<{ block: LoweredParagraph }> };
      expect(output.measured[0]!.block.runs[0]).not.toHaveProperty('color');
      expect(output.measured[0]!.block.runs[0]).not.toHaveProperty('underline');
      const blocks = session.yrsBlocksForStory('body', env) as LoweredParagraph[];
      expect(blocks[0]!.runs[0]).not.toHaveProperty('color');
      expect(blocks[0]!.runs[0]).not.toHaveProperty('underline');
      expect(blocks[0]!.runs[0]?.hyperlink).toMatchObject({ href: '#heading', noDefaultStyle: true });
      expect(blocks[0]!.runs[1]?.color).toBe('#112233');
      const firstPara = session.paragraphs('body')[0]!.paraId;
      const firstRange = { story: 'body', start: { paraId: firstPara, offset: 0 }, end: { paraId: firstPara, offset: 9 } };
      session.applyParagraphStyle(firstRange, 'Body');
      expect((session.yrsBlocksForStory('body', env) as LoweredParagraph[])[0]!.runs[0]?.color).toBe('#0000FF');
      session.applyParagraphStyle(firstRange, '10');
      expect((session.yrsBlocksForStory('body', env) as LoweredParagraph[])[0]!.runs[0]).not.toHaveProperty('color');
      const first = session.storySegments('body')[0]!;
      expect(first.attributes.textColor).toMatchObject({ inheritedHyperlink: true });
      expect(first.attributes.underline).toMatchObject({ inheritedHyperlink: true });
      const saved = await parseDocx(await repackDocx(yrsToDocument(session, parsed)), { preloadFonts: false });
      const runs = linkedRuns(saved);
      expect(runs[0]?.formatting?.styleId).toBe('af');
      expect(runs[0]?.formatting?.color).toBeUndefined();
      expect(runs[0]?.formatting?.underline).toBeUndefined();
      expect(runs[1]?.formatting?.color?.rgb).toBe('112233');
      expect(runs[1]?.formatting?.underline?.style).toBe('double');
      expect(runs[2]?.formatting?.color?.rgb).toBe('0000FF');
      expect(runs[2]?.formatting?.underline?.style).toBe('single');
      documentToYrs(reopened, saved);
      expect(reopened.storySegments('body')[0]!.attributes).toEqual(first.attributes);

      const paraId = session.paragraphs('body')[0]!.paraId;
      session.formatRange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 9 } },
        { color: { rgb: '0000FF' }, underline: { style: 'single' } });
      expect(session.storySegments('body')[0]!.attributes.textColor).not.toHaveProperty('inheritedHyperlink');
      expect(session.storySegments('body')[0]!.attributes.underline).not.toHaveProperty('inheritedHyperlink');
      const edited = await parseDocx(await repackDocx(yrsToDocument(session, parsed)), { preloadFonts: false });
      const editedRun = linkedRuns(edited)[0]!;
      expect(editedRun.formatting?.color?.rgb).toBe('0000FF');
      expect(editedRun.formatting?.underline?.style).toBe('single');
      session.setHyperlink({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 9 } },
        { href: '#changed' });
      expect(session.storySegments('body')[0]!.attributes.hyperlink).toMatchObject({
        href: '#changed',
      });
      const relinked = await parseDocx(await repackDocx(yrsToDocument(session, parsed)), { preloadFonts: false });
      const paragraph = relinked.package.document.content[0];
      if (paragraph?.type !== 'paragraph') throw new Error('missing relinked paragraph');
      const link = paragraph.content[0];
      if (link?.type !== 'hyperlink') throw new Error('missing changed hyperlink');
      expect(link.anchor).toBe('changed');
      const run = link.children[0];
      if (run?.type !== 'run') throw new Error('missing changed run');
      expect(run.formatting?.color?.rgb).toBe('0000FF');
      expect(run.formatting?.underline?.style).toBe('single');
    } finally {
      session.destroy();
      reopened.destroy();
    }
  });
}

test('TOC names cannot turn headings or character styles into TOC paragraphs', () => {
  const document = { package: { document: { content: [] }, styles: { styles: [
    { styleId: '10', type: 'paragraph', name: 'toc 1' },
    { styleId: 'TOC2', type: 'paragraph' },
    { styleId: 'Heading', type: 'paragraph', name: 'TOC Heading' },
    { styleId: 'Character', type: 'character', name: 'toc 3' },
    { styleId: '', type: 'paragraph', name: 'toc 4' },
  ] } } } as Document;
  const request = buildResidentRegionLayoutRequest(document, 24, { tocStyleIds: ['CustomTOC'] });
  expect(request.renderEnv.tocStyleIds).toEqual(['CustomTOC', '10', 'TOC2']);
});

test('persisted and peer TOC state preserves formatting without hyperlink provenance', async () => {
  const parsed = await parseDocx(source().buffer, { preloadFonts: false });
  const env = buildResidentRegionLayoutRequest(parsed, 24, {}).renderEnv;
  for (const style of ['10', 'TOC1']) {
    for (const delivery of ['persisted', 'peer']) {
      for (const inheritedColor of [false, true]) {
        const origin = await createYrsSession({ clientId: 74510 });
        const received = await createYrsSession({ clientId: 74511 });
        const reopened = await createYrsSession({ clientId: 74512 });
        try {
          documentToYrs(origin, parsed);
          origin.setParagraphAttr(origin.paragraphs('body')[0]!.paraId, 'pStyle', style);
          origin.applyRawOps('body', [{ op: 'format', index: 0, len: 9, attrs: {
            hyperlink: { href: '#heading' },
            textColor: { rgb: '112233', ...(inheritedColor ? { inheritedHyperlink: true } : {}) },
            underline: { style: 'double' },
          } }]);
          if (delivery === 'persisted') received.loadState(origin.encodeState());
          else received.applyUpdate(origin.encodeStateAsUpdate());
          const before = received.encodeState();
          const live = (received.yrsBlocksForStory('body', env) as LoweredParagraph[])[0]!.runs[0]!;
          expect(live.color).toBe(inheritedColor ? undefined : '#112233');
          expect(live.underline).toMatchObject({ style: 'double' });
          expect(live.hyperlink).toMatchObject({ href: '#heading', noDefaultStyle: true });
          expect(received.encodeState()).toEqual(before);
          const saved = await parseDocx(await repackDocx(yrsToDocument(received, parsed)), { preloadFonts: false });
          const run = linkedRuns(saved)[0]!;
          expect(run.formatting?.color?.rgb).toBe(inheritedColor ? undefined : '112233');
          expect(run.formatting?.underline?.style).toBe('double');
          documentToYrs(reopened, saved);
          const after = (reopened.yrsBlocksForStory('body', env) as LoweredParagraph[])[0]!.runs[0]!;
          expect(after.color).toBe(live.color);
          expect(after.underline).toEqual(live.underline);
        } finally {
          origin.destroy();
          received.destroy();
          reopened.destroy();
        }
      }
    }
  }
});
