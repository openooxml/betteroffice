import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, test } from 'bun:test';
import JSZip from 'jszip';

import {
  exportDocxMarkdown,
  exportDocxStructured,
  findDocxContentControls,
  listDocxContentControls,
  type DocxContentControlsSnapshot,
  type DocxExportOptions,
  type DocxMarkdownContent,
  type DocxStructuredContent,
} from '../../packages/docx/src/core';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import { expectSavedPartsEqual } from '../native-parity';
import { openDocument, renderDocxMarkdown } from './index.js';

const fixture = readFileSync(
  new URL('../../apps/demo/public/betteroffice-demo.docx', import.meta.url)
);
const layoutInput = JSON.parse(
  readFileSync(
    new URL(
      '../../crates/docx-layout/tests/fixtures/single-page-multi-paragraph.input.json',
      import.meta.url
    ),
    'utf8'
  )
);
const font = readFileSync(
  new URL('../../crates/docx-raster/tests/assets/Carlito-Regular.ttf', import.meta.url)
);
const exportFixture = readFileSync(
  new URL('../../crates/docx-edit/tests/fixtures/structured-export/principal.docx', import.meta.url)
);
const controlsFixture = readFileSync(
  new URL('../../packages/docx/src/yrs/__fixtures__/content-controls/template.docx', import.meta.url)
);

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)
      )
    )
  )
);

describe('@betteroffice/docx-native', () => {
  test('opens, inspects, lays out, renders, and saves a document', async () => {
    const document = await openDocument(fixture, { author: 'Node test' });

    const structure = await document.structure;
    const paragraphIds = await document.paragraphIds;
    const text = await document.text;

    expect(structure.bodyParagraphs).toBeGreaterThan(0);
    expect(document.author).toBe('Node test');
    expect(paragraphIds).toHaveLength(structure.bodyParagraphs);
    expect(text.length).toBeGreaterThan(0);
    const layout = await document.layout(layoutInput);
    await document.registerFont({ family: 'Calibri', data: font });
    const rendered = await document.renderPage(layout.displayList);

    expect(rendered.data.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    expect((await document.save()).subarray(0, 2)).toEqual(Buffer.from('PK'));
  });

  test('rejects invalid edit origins', async () => {
    expect(() => openDocument(fixture, { origin: 'unknown' })).toThrow('origin must be');
  });

  test('applies optional parse limits and preserves unspecified defaults', async () => {
    const document = await openDocument(exportFixture, { limits: { maxParagraphs: 500_000 } });
    expect((await document.structure).bodyParagraphs).toBeGreaterThan(1);
    const defaults = await openDocument(exportFixture, { limits: {} });
    expect(await document.text).toBe(await defaults.text);
    await expect(openDocument(exportFixture, { limits: { maxParagraphs: 1 } })).rejects.toThrow();
    await expect(openDocument(exportFixture, { limits: { maxXmlBytes: 1 } })).rejects.toThrow();
    for (const maxBlocks of [-1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1]) {
      expect(() => openDocument(exportFixture, { limits: { maxBlocks } })).toThrow(
        'maxBlocks must be a nonnegative safe integer'
      );
    }
  });

  test('exports structured content and Markdown without a session envelope', async () => {
    const document = await openDocument(exportFixture);
    const content = (await document.exportStructured({
      revisionView: 'markup',
      stories: ['body', 'headers'],
      includeFormatting: false,
    })) as DocxStructuredContent;
    expect(content.schemaVersion).toBe(1);
    expect(content.anchorScope).toBe('snapshot');
    expect(content.revisionView).toBe('markup');
    expect(content.includeFormatting).toBe(false);
    expect(content.includedStories).toEqual(['body', 'headers']);
    expect(content.stories[0].blocks.length).toBeGreaterThan(1);
    expect(content).not.toHaveProperty('ok');
    expect(content).not.toHaveProperty('version');
    const markdown = (await document.exportMarkdown({
      revisionView: 'markup',
      stories: ['body', 'headers'],
      includeFormatting: false,
    })) as DocxMarkdownContent;
    expect(markdown.markdown).toContain('Structured export');
    expect(markdown.anchors.length).toBeGreaterThan(0);
    const rendering = renderDocxMarkdown(content);
    expect(rendering).toBeInstanceOf(Promise);
    expect(await rendering).toBe(markdown.markdown);
    expect(await renderDocxMarkdown(content, { maxBytes: 8_388_608 })).toBe(markdown.markdown);
    const truncated = (await document.exportStructured({
      revisionView: 'accepted',
      maxBlocks: 1,
    })) as DocxStructuredContent;
    expect(truncated.truncated).toBe(true);

    await document.replaceParagraphText('00000001', 'Updated native export');
    const updated = (await document.exportMarkdown({ revisionView: 'accepted' })) as DocxMarkdownContent;
    expect(updated.markdown).toContain('Updated native export');
  });

  test('rejects export refusals with the failure message and data', async () => {
    const document = await openDocument(exportFixture);
    for (const operation of [
      () => document.exportStructured({ revisionView: 'accepted', maxBlocks: 2_000_000 }),
      () => document.exportMarkdown({ revisionView: 'accepted', maxBytes: 67_108_865 }),
    ]) {
      const refused = await operation().catch((error) => error);
      expect(refused).toBeInstanceOf(Error);
      expect(refused.name).toBe('DocxExportError');
      expect(refused.failure).toEqual({
        code: 'limit-exceeded',
        target: null,
        message: expect.any(String),
      });
      expect(refused.message).toBe(refused.failure.message);
    }
    const content = await document.exportStructured({ revisionView: 'accepted' });
    const rendering = renderDocxMarkdown(content, { maxBytes: 16 });
    expect(rendering).toBeInstanceOf(Promise);
    const refused = await rendering.catch((error) => error);
    expect(refused).toBeInstanceOf(Error);
    expect(refused).toMatchObject({
      name: 'DocxExportError',
      failure: { code: 'invalid-options', target: null },
    });
    expect(refused.message).toBe(refused.failure.message);
  });

  test('rejects malformed Markdown content and options asynchronously', async () => {
    const malformedContent = renderDocxMarkdown({ schemaVersion: 2 });
    expect(malformedContent).toBeInstanceOf(Promise);
    await expect(malformedContent).rejects.toThrow();
    const document = await openDocument(exportFixture);
    const content = await document.exportStructured({ revisionView: 'accepted' });
    const malformedOptions = renderDocxMarkdown(content, { maxBytes: -1 });
    expect(malformedOptions).toBeInstanceOf(Promise);
    await expect(malformedOptions).rejects.toThrow();
  });

  test('lists and finds content controls with optional story filters', async () => {
    const document = await openDocument(controlsFixture);
    const snapshot = (await document.listContentControls()) as DocxContentControlsSnapshot;
    expect(snapshot.schemaVersion).toBe(1);
    expect(snapshot.anchorScope).toBe('snapshot');
    expect(snapshot.complete).toBe(true);
    expect(snapshot.controls.length).toBeGreaterThan(2);
    expect(snapshot).not.toHaveProperty('version');
    const duplicates = (await document.findContentControls({
      kind: 'tag',
      tag: 'account.reference',
    })) as DocxContentControlsSnapshot;
    expect(duplicates.controls).toHaveLength(2);
    const address = (await document.findContentControls({
      kind: 'alias',
      alias: 'Address',
    })) as DocxContentControlsSnapshot;
    expect(address.controls).toHaveLength(1);
    const byId = (await document.findContentControls({
      kind: 'id',
      controlId: address.controls[0].controlId,
    })) as DocxContentControlsSnapshot;
    expect(byId.controls).toEqual(address.controls);
    const byOoxmlId = (await document.findContentControls({
      kind: 'ooxmlId',
      ooxmlId: '105',
    })) as DocxContentControlsSnapshot;
    expect(byOoxmlId.controls).toEqual(address.controls);
    const missing = (await document.findContentControls({
      kind: 'tag',
      tag: 'ACCOUNT.reference',
    })) as DocxContentControlsSnapshot;
    expect(missing.controls).toEqual([]);
    const options = { stories: ['headers'] as const };
    const headers = (await document.listContentControls(options)) as DocxContentControlsSnapshot;
    expect(headers.controls.map((control) => control.tag)).toEqual(['document.title']);
    const title = (await document.findContentControls(
      { kind: 'tag', tag: 'document.title' },
      options
    )) as DocxContentControlsSnapshot;
    expect(title.controls).toEqual(headers.controls);
  });

  test('rejects content-control refusals with the failure message and data', async () => {
    const document = await openDocument(controlsFixture);
    for (const operation of [
      () => document.listContentControls({ maxControls: 2 }),
      () => document.findContentControls({ kind: 'tag', tag: 'account.reference' }, { maxControls: 1 }),
    ]) {
      const refused = await operation().catch((error) => error);
      expect(refused).toBeInstanceOf(Error);
      expect(refused.name).toBe('DocxContentControlsError');
      expect(refused.failure.code).toBe('limit-exceeded');
      expect(refused.message).toBe(refused.failure.message);
    }
    const invalid = await document.listContentControls({ maxControls: 0 }).catch((error) => error);
    expect(invalid.failure.code).toBe('invalid-options');
  });

  test('matches the WASM bytes-level exports and content controls', async () => {
    for (const bytes of [exportFixture, controlsFixture]) {
      const document = await openDocument(bytes);
      for (const revisionView of ['accepted', 'original', 'markup'] as const) {
        const options: DocxExportOptions = {
          revisionView,
          stories: ['body', 'headers', 'footers', 'footnotes', 'endnotes', 'comments'],
          includeFormatting: false,
          maxBlocks: 10_000,
          maxBytes: 8_388_608,
        };
        expect(await document.exportStructured(options)).toEqual(
          await exportDocxStructured(bytes, options)
        );
        expect(await document.exportMarkdown(options)).toEqual(await exportDocxMarkdown(bytes, options));
      }
      expect(await document.listContentControls()).toEqual(await listDocxContentControls(bytes));
      const options = { stories: ['body', 'headers'] as const, maxControls: 10_000, maxBytes: 8_388_608 };
      expect(await document.listContentControls(options)).toEqual(
        await listDocxContentControls(bytes, options)
      );
      const query = { kind: 'tag', tag: 'account.reference' } as const;
      expect(await document.findContentControls(query, options)).toEqual(
        await findDocxContentControls(bytes, query, options)
      );
    }
  });

  test('accepts now as a save timestamp alias and rejects conflicting values', async () => {
    const document = await openDocument(fixture, { timestamp: '2000-01-01T00:00:00.000Z' });
    const now = '2026-05-06T07:08:09.000Z';
    const options = { updateModifiedDate: true, modifiedBy: 'Native alias test' };
    const timestamp = await document.save({ ...options, timestamp: now });
    const aliased = await document.save({ ...options, now });
    await expectSavedPartsEqual(timestamp, aliased);
    await expectSavedPartsEqual(timestamp, await document.save({ ...options, timestamp: now, now }));
    const zip = await JSZip.loadAsync(aliased);
    const properties = await zip.file('docProps/core.xml')!.async('string');
    expect(properties).toContain(now);
    expect(properties).toContain('Native alias test');
    const conflict = document.save({ timestamp: document.timestamp, now });
    expect(conflict).toBeInstanceOf(Promise);
    await expect(conflict).rejects.toBeInstanceOf(TypeError);
    await expect(conflict).rejects.toThrow('now and timestamp must match when both are provided');
    expect(document.timestamp).toBe('2000-01-01T00:00:00.000Z');
    await expectSavedPartsEqual(timestamp, await document.save({ ...options, timestamp: now }));
  });
});
