import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { assertBookmarkInvariant, collectBookmarkBoundaries } from '../docx/__testSupport__/bookmarkInvariant';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { BlockContent, Document, Paragraph, Shape } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import {
  blockBookmarkFixture, blockContainers, bookmarkBlocks, bookmarkEnd,
  bookmarkParagraph, bookmarkPart, bookmarkStart, drawingBookmarkFixture, type BlockContainer,
} from './__fixtures__/blockBookmarks';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';
import { restoreBlockBookmarks, yrsToDocument } from './yrsToDocument';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

function textBox(document: Document): Shape {
  const host = document.package.document.content.find(block => block.type === 'paragraph' && block.paraId === '00000020');
  if (host?.type !== 'paragraph') throw new Error('Missing text box paragraph');
  for (const child of host.content) {
    if (child.type !== 'run') continue;
    for (const item of child.content) if (item.type === 'shape') return item.shape;
  }
  throw new Error('Missing text box');
}

function containerBlocks(document: Document, container: BlockContainer): BlockContent[] {
  const pkg = document.package;
  switch (container) {
    case 'cell': {
      const table = pkg.document.content.find(block => block.type === 'table');
      if (table?.type !== 'table') throw new Error('Missing table');
      return table.rows[0]!.cells[0]!.content;
    }
    case 'control': {
      const control = pkg.document.content.find(block => block.type === 'blockSdt');
      if (control?.type !== 'blockSdt') throw new Error('Missing control');
      return control.content;
    }
    case 'header': return [...pkg.headers!.values()][0]!.content;
    case 'footer': return [...pkg.footers!.values()][0]!.content;
    case 'footnote': return pkg.footnotes!.find(note => note.id === 1)!.content;
    case 'endnote': return pkg.endnotes!.find(note => note.id === 1)!.content;
    case 'comment': return pkg.document.comments![0]!.blockContent!;
    case 'textbox': return textBox(document).textBody!.content as BlockContent[];
  }
}

function containerXml(bytes: Uint8Array, container: BlockContainer): string {
  const xml = new TextDecoder().decode(unzipContainer(bytes)[bookmarkPart(container)]);
  const root = {
    cell: 'tc', control: 'sdtContent', header: 'hdr', footer: 'ftr',
    footnote: 'footnote', endnote: 'endnote', comment: 'comment', textbox: 'txbxContent',
  }[container];
  const start = new RegExp(`<w:${root}(?=[\\s>])[^>]*>`).exec(xml);
  expect(start).not.toBeNull();
  return xml.slice(start!.index + start![0].length, xml.indexOf(`</w:${root}>`, start!.index));
}

function assertPair(source: Uint8Array, bytes: Uint8Array, container: BlockContainer, order: string[]): void {
  assertBookmarkInvariant(source, bytes);
  const ids = collectBookmarkBoundaries(bytes).get(bookmarkPart(container));
  expect(ids?.get('7')).toEqual({ Start: 1, End: 1 });
  const xml = containerXml(bytes, container);
  expect(xml.match(/<w:bookmarkStart\b/g)).toHaveLength(1);
  expect(xml.match(/<w:bookmarkEnd\b/g)).toHaveLength(1);
  let offset = 0;
  for (const fragment of order) {
    const position = xml.indexOf(fragment, offset);
    expect(position).toBeGreaterThanOrEqual(offset);
    offset = position + fragment.length;
  }
}

for (const container of blockContainers) {
  for (const edited of [false, true]) {
    it(`${container} satisfies the bookmark invariant on both save paths ${edited ? 'after insertText' : 'untouched'}`, async () => {
      const source = blockBookmarkFixture(container);
      const base = await parseDocx(source.buffer, { preloadFonts: false });
      if (edited) {
        const first = base.package.document.content[0] as Paragraph;
        first.content.unshift({ type: 'run', content: [{ type: 'text', text: 'Edited ' }] });
      }
      const outputs = [new Uint8Array(await repackDocx(base))];
      const session = await createYrsSession();
      try {
        session.openDocx(source, true);
        if (edited) session.insertText({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 }, 'Edited ');
        outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
      } finally { session.destroy(); }
      for (const saved of outputs) {
        assertPair(source, saved, container, ['>Before<', bookmarkStart, '>Anchor<', bookmarkEnd, '>After<']);
        if (edited) expect(new TextDecoder().decode(unzipContainer(saved)['word/document.xml'])).toContain('Edited');
      }
    });
  }
}

for (const kind of ['choice', 'group', 'vml'] as const) {
  for (const edited of [false, true]) {
    it(`keeps ${kind} text box boundaries on both save paths ${edited ? 'after a host edit' : 'untouched'}`, async () => {
      const source = drawingBookmarkFixture(kind);
      const base = await parseDocx(source.buffer, { preloadFonts: false });
      if (edited) {
        const host = base.package.document.content.find(block => block.type === 'paragraph' && block.paraId === '00000020') as Paragraph;
        host.content.unshift({ type: 'run', content: [{ type: 'text', text: 'Edited ' }] });
      }
      const outputs = [new Uint8Array(await repackDocx(base))];
      const session = await createYrsSession();
      try {
        session.openDocx(source, true);
        if (edited) session.insertText({ story: 'body', paraId: '00000020', offset: 0 }, 'Edited ');
        outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
      } finally { session.destroy(); }
      for (const saved of outputs) {
        assertBookmarkInvariant(source, saved);
        expect(collectBookmarkBoundaries(saved)).toEqual(collectBookmarkBoundaries(source));
        const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
        expect(xml).toContain(bookmarkStart.replace("'&#55;'", "'8'"));
        expect(xml).toContain(bookmarkEnd.replace('"7"', '"8"'));
        if (kind !== 'vml') {
          expect(xml).toContain(bookmarkStart);
          expect(xml).toContain(bookmarkEnd);
          expect(xml).toContain('<mc:Fallback>');
        }
        if (edited) expect(xml).toContain('Edited');
      }
    });
  }
}

for (const container of ['footnote', 'endnote'] as const) {
  it(`keeps untouched ${container} bytes with block bookmarks, revision ids and field runs`, async () => {
    const initial = blockBookmarkFixture(container, bookmarkParagraph('00000002', 'Anchor'));
    const baseline = await repackDocx(await parseDocx(initial.buffer, { preloadFonts: false }));
    const parts = new Map(Object.entries(unzipContainer(new Uint8Array(baseline))));
    const path = bookmarkPart(container);
    const note = `<w:${container} w:id="1">${bookmarkStart}<w:p w14:paraId="00000002" w:rsidR="00112233" w:rsidRDefault="44556677" w:rsidP="8899AABB"><w:r w:rsidR="00112233"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.test/note" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Anchor</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>${bookmarkEnd}</w:${container}>`;
    const xml = new TextDecoder().decode(parts.get(path));
    const pattern = new RegExp(`<w:${container}(?=[\\s>])[^>]*>[\\s\\S]*?</w:${container}>`);
    expect(pattern.test(xml)).toBe(true);
    parts.set(path, toBytes(xml.replace(pattern, () => note)));
    const source = new Uint8Array(rezipPartsToArrayBuffer(parts));
    const base = await parseDocx(source.buffer, { preloadFonts: false });
    const notes = container === 'footnote' ? base.package.footnotes : base.package.endnotes;
    expect(notes![0]!.verbatimXml).toBe(note);
    const session = await createYrsSession();
    try {
      session.openDocx(source, true);
      for (const edited of [false, true]) {
        if (edited) {
          (base.package.document.content[0] as Paragraph).content.unshift({ type: 'run', content: [{ type: 'text', text: 'Edited ' }] });
          session.insertText({ story: 'body', paraId: '00000010', offset: 0 }, 'Edited ');
        }
        const outputs = [new Uint8Array(await repackDocx(base)), new Uint8Array((await saveYrsDocx(session)).bytes)];
        for (const saved of outputs) {
          assertBookmarkInvariant(source, saved);
          expect(unzipContainer(saved)[path]).toEqual(parts.get(path)!);
        }
      }
    } finally { session.destroy(); }
  });
}

it('reports every missing boundary by part and decoded id', () => {
  const parts = unzipContainer(blockBookmarkFixture('cell'));
  parts['word/header1.xml'] = toBytes('<b:hdr xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><b:bookmarkStart b:id="&#56;" b:name="Header"/><b:bookmarkEnd b:id="8"/></b:hdr>');
  const source = new Uint8Array(rezipPartsToArrayBuffer(new Map(Object.entries(parts))));
  const xml = new TextDecoder().decode(parts['word/document.xml']);
  parts['word/document.xml'] = toBytes(xml.replace(bookmarkStart, '').replace(bookmarkEnd, ''));
  delete parts['word/header1.xml'];
  const saved = new Uint8Array(rezipPartsToArrayBuffer(new Map(Object.entries(parts))));
  let failure = '';
  try { assertBookmarkInvariant(source, saved); }
  catch (error) { failure = (error as Error).message; }
  for (const part of ['word/document.xml: w:id=7', 'word/header1.xml: w:id=8']) {
    expect(failure).toContain(`${part} missing Start (1)`);
    expect(failure).toContain(`${part} missing End (1)`);
  }
});

it('rejects duplicate source boundaries instead of skipping them', () => {
  const start: BlockContent = { type: 'rawXml', xml: bookmarkStart };
  expect(() => restoreBlockBookmarks([], [start, start], 'body')).toThrow('Duplicate block bookmark start for id 7');
});

type Scenario = 'delete' | 'merge' | 'move' | 'tail' | 'empty' | 'start' | 'end' | 'both';
const scenarios: Scenario[] = ['delete', 'merge', 'move', 'tail', 'empty', 'start', 'end', 'both'];

function storyFor(session: YrsSession): string {
  const story = session.storyIds().find(story => session.paragraphs(story).some(paragraph => paragraph.text === 'Anchor'));
  if (!story) throw new Error('Missing bookmark story');
  return story;
}

function editStory(session: YrsSession, story: string, scenario: Scenario, native = true): void {
  const anchor = session.paragraphs(story).find(paragraph => paragraph.text === 'Anchor')!;
  const span = session.locateParagraph(story, anchor.paraId);
  if (scenario === 'merge') {
    session.mergeParagraphs(story, session.paragraphs(story)[0]!.paraId);
  } else if (scenario === 'delete') {
    if (native) {
      expect(session.applyEdits({
        expectVersion: session.version(),
        steps: [{ op: 'deleteParagraphs', story, firstParaId: anchor.paraId, lastParaId: anchor.paraId }],
      }).ok).toBe(true);
    } else {
      session.applyRawOps(story, [{ op: 'delete', index: span.start, len: span.end - span.start + 1 }]);
    }
    expect(session.paragraphs(story).some(paragraph => paragraph.paraId === anchor.paraId)).toBe(false);
  } else if (scenario === 'move') {
    const segment = session.storySegments(story).find(segment => segment.kind === 'pilcrow' && segment.paraId === anchor.paraId)!;
    if (segment.kind !== 'pilcrow') throw new Error('Missing anchor paragraph mark');
    session.applySeedRawOps(story, [
      { op: 'delete', index: span.start, len: span.end - span.start + 1 },
      { op: 'insert', index: 13, text: 'Anchor' },
      { op: 'insertEmbed', index: 19, kind: 'pilcrow', payload: { ...segment.properties, paraId: anchor.paraId } },
    ]);
    expect(session.paragraphs(story).map(paragraph => paragraph.text)).toEqual(['Before', 'After', 'Anchor']);
  } else if (scenario === 'start' || scenario === 'end' || scenario === 'both') {
    const kinds = scenario === 'both' ? ['start', 'end'] : [scenario];
    session.applyRawOps(story, [{
      op: 'setEmbedAttr', index: span.end, key: 'bookmarks',
      value: kinds.map(kind => ({ id: 7, kind, name: 'Block&Range', offset: kind === 'start' ? 0 : 6 })),
    }]);
  } else {
    session.applyRawOps(story, [{ op: 'delete', index: span.start, len: span.end - span.start + 1 }]);
  }
}

function editBlocks(blocks: BlockContent[], scenario: Scenario): BlockContent[] {
  const paragraphs = structuredClone(blocks.filter((block): block is Paragraph => block.type === 'paragraph'));
  const index = paragraphs.findIndex(paragraph => paragraph.paraId === '00000002');
  const anchor = paragraphs[index]!;
  if (scenario === 'merge') {
    paragraphs[index - 1]!.content.push(...anchor.content);
    paragraphs.splice(index, 1);
  } else if (scenario === 'move') {
    paragraphs.splice(index, 1);
    paragraphs.push(anchor);
  } else if (scenario === 'start' || scenario === 'end' || scenario === 'both') {
    if (scenario !== 'end') anchor.content.unshift({ type: 'bookmarkStart', id: 7, name: 'Block&Range' });
    if (scenario !== 'start') anchor.content.push({ type: 'bookmarkEnd', id: 7 });
  } else {
    paragraphs.splice(index, 1);
  }
  return paragraphs;
}

function editStoredBlocks(session: YrsSession, base: Document, container: 'comment' | 'textbox', scenario: Scenario): void {
  const blocks = editBlocks(containerBlocks(base, container), scenario);
  if (container === 'comment') {
    const ranges = session.resolveComment('1').map(({ start, end }) => [start, end] as const);
    session.applyRawOps('body', [{ op: 'setComment', id: '1', ranges, body: blocks }]);
  } else {
    const shape = structuredClone(textBox(base));
    shape.textBody!.content = blocks as Paragraph[];
    delete shape.drawingXml;
    delete shape.drawingRunXml;
    delete shape.drawingModel;
    let index = 0;
    for (const segment of session.storySegments('body')) {
      if (segment.kind === 'embed' && segment.embedKind === 'shape') {
        session.applyRawOps('body', [{ op: 'setEmbedAttr', index, key: 'shapeJson', value: JSON.stringify(shape) }]);
        return;
      }
      index += segment.kind === 'text' ? segment.text.length : 1;
    }
    throw new Error('Missing shape embed');
  }
}

for (const seeder of ['native', 'projected']) {
  for (const kind of ['choice', 'group'] as const) {
    for (const scenario of ['delete', 'merge', 'move', 'start', 'end'] as const) {
      it(`${seeder} keeps ${kind} text box sources after ${scenario}`, async () => {
        const source = drawingBookmarkFixture(kind);
        const base = await parseDocx(source.buffer, { preloadFonts: false });
        const session = await createYrsSession();
        try {
          if (seeder === 'native') session.openDocx(source, true);
          else documentToYrs(session, base);
          const shape = structuredClone(textBox(base));
          const box = kind === 'group' ? shape.children![0]! : shape;
          box.textBody!.content = editBlocks(box.textBody!.content as BlockContent[], scenario) as Paragraph[];
          let index = 0;
          let changed = false;
          for (const segment of session.storySegments('body')) {
            if (segment.kind === 'embed' && segment.embedKind === 'shape') {
              session.applyRawOps('body', [{ op: 'setEmbedAttr', index, key: 'shapeJson', value: JSON.stringify(shape) }]);
              changed = true;
              break;
            }
            index += segment.kind === 'text' ? segment.text.length : 1;
          }
          expect(changed).toBe(true);
          const outputs = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
          if (seeder === 'native') outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
          for (const saved of outputs) {
            assertPair(source, saved, 'textbox', expectedOrder(scenario));
            expect(collectBookmarkBoundaries(saved)).toEqual(collectBookmarkBoundaries(source));
            expect(new TextDecoder().decode(unzipContainer(saved)['word/document.xml'])).toContain('<mc:Fallback>');
          }
        } finally { session.destroy(); }
      });
    }
  }
}

function expectedOrder(scenario: Scenario): string[] {
  if (scenario === 'move') return ['>Before<', '>After<', '<w:bookmarkStart', '<w:bookmarkEnd', '>Anchor<'];
  if (scenario === 'merge') return ['Before', 'Anchor', '<w:bookmarkStart', '<w:bookmarkEnd', '>After<'];
  if (scenario === 'delete') return ['>Before<', '<w:bookmarkStart', '<w:bookmarkEnd', '>After<'];
  if (scenario === 'tail') return ['>Before<', '<w:bookmarkStart', '<w:bookmarkEnd'];
  if (scenario === 'empty') return ['<w:bookmarkStart', '<w:bookmarkEnd'];
  return ['>Before<', '<w:bookmarkStart', '>Anchor<', '<w:bookmarkEnd', '>After<'];
}

for (const seeder of ['native', 'projected']) {
  for (const container of blockContainers) {
    for (const scenario of scenarios) {
      it(`${seeder} keeps ${container} boundaries after ${scenario}`, async () => {
        const blocks = scenario === 'empty' ? `${bookmarkStart}${bookmarkParagraph('00000002', 'Anchor')}${bookmarkEnd}`
          : scenario === 'tail' ? bookmarkBlocks.replace(bookmarkParagraph('00000003', 'After'), '') : bookmarkBlocks;
        const source = blockBookmarkFixture(container, blocks);
        const base = await parseDocx(source.buffer, { preloadFonts: false });
        const session = await createYrsSession();
        try {
          if (seeder === 'native') session.openDocx(source, true);
          else documentToYrs(session, base);
          if (container === 'comment' || container === 'textbox') editStoredBlocks(session, base, container, scenario);
          else editStory(session, storyFor(session), scenario, seeder === 'native');
          for (let round = 0; round < 2; round += 1) {
            const outputs = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
            if (seeder === 'native') outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
            for (const saved of outputs) {
              assertPair(source, saved, container, expectedOrder(scenario));
              if (container === 'cell' && scenario === 'empty') {
                const xml = containerXml(saved, container);
                expect(xml.indexOf('<w:bookmarkEnd')).toBeLessThan(xml.lastIndexOf('<w:p'));
              }
              const reopened = await parseDocx(saved.slice().buffer, { preloadFonts: false });
              expect(containerBlocks(reopened, container)).toBeDefined();
            }
          }
        } finally { session.destroy(); }
      });
    }
  }
}

for (const seeder of ['native', 'projected']) {
  for (const container of blockContainers.filter(container => container !== 'comment' && container !== 'textbox')) {
    for (const ids of ['missing', 'repeated']) {
      it(`${seeder} keeps ${container} anchors with ${ids} paragraph ids`, async () => {
        const blocks = bookmarkBlocks.replace(/ w14:paraId="[^"]*"/g, ids === 'missing' ? '' : ' w14:paraId="12345678"');
        const source = blockBookmarkFixture(container, blocks);
        const base = await parseDocx(source.buffer, { preloadFonts: false });
        const session = await createYrsSession();
        try {
          if (seeder === 'native') session.openDocx(source, true);
          else documentToYrs(session, base);
          editStory(session, storyFor(session), 'move');
          const outputs = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
          if (seeder === 'native') outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
          for (const saved of outputs) assertPair(source, saved, container, expectedOrder('move'));
        } finally { session.destroy(); }
      });
    }
  }
}

for (const body of ['text', 'blocks'] as const) {
  it(`rejects replacing a nested bookmark comment with ${body}`, async () => {
    const source = blockBookmarkFixture('comment',
      `${bookmarkStart}<w:sdt><w:sdtContent>${bookmarkParagraph('00000002', 'Nested')}${bookmarkEnd}</w:sdtContent></w:sdt>${bookmarkParagraph('00000003', 'After')}`
    );
    const base = await parseDocx(source.buffer, { preloadFonts: false });
    expect(base.package.document.comments![0]!.blockContent![1]!.type).toBe('blockSdt');
    const session = await createYrsSession();
    try {
      session.openDocx(source, true);
      const ranges = session.resolveComment('1').map(({ start, end }) => [start, end] as const);
      const replacement = body === 'text' ? 'Changed'
        : [{ type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Changed' }] }] }];
      session.applyRawOps('body', [{ op: 'setComment', id: '1', ranges, body: replacement }]);
      if (body === 'blocks') {
        expect(() => yrsToDocument(session, base)).toThrow('Cannot replace comment body with bookmarks in nested blocks');
      }
      await expect(saveYrsDocx(session)).rejects.toThrow('Cannot replace comment body with bookmarks in nested blocks');
    } finally { session.destroy(); }
  });
}

function aliasedBookmarkFixture(container: 'cell' | 'comment'): Uint8Array<ArrayBuffer> {
  const start = bookmarkStart.replace('w:id', 'b:id').replace('w:name', 'b:name');
  const end = bookmarkEnd.replace('w:id', 'b:id');
  const source = blockBookmarkFixture(container, bookmarkBlocks.replace(bookmarkStart, start).replace(bookmarkEnd, end));
  const parts = new Map(Object.entries(unzipContainer(source)));
  const path = bookmarkPart(container);
  const xml = new TextDecoder().decode(parts.get(path));
  const root = container === 'cell' ? '<w:document ' : '<w:comments ';
  parts.set(path, toBytes(xml.replace(root, `${root}xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `)));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

for (const seeder of ['native', 'projected']) {
  for (const container of ['cell', 'comment'] as const) {
    for (const scenario of ['move', 'start', 'end', 'both'] as const) {
      it(`${seeder} pairs aliased ${container} boundaries after ${scenario}`, async () => {
        const source = aliasedBookmarkFixture(container);
        const base = await parseDocx(source.buffer, { preloadFonts: false });
        const session = await createYrsSession();
        try {
          if (seeder === 'native') session.openDocx(source, true);
          else documentToYrs(session, base);
          if (container === 'comment') editStoredBlocks(session, base, container, scenario);
          else editStory(session, storyFor(session), scenario, seeder === 'native');
          const outputs = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
          if (seeder === 'native') outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
          for (const saved of outputs) assertPair(source, saved, container, expectedOrder(scenario));
        } finally { session.destroy(); }
      });
    }
  }
}

it('rejects an unbound bookmark identity prefix instead of treating the endpoints as anonymous', () => {
  const base: BlockContent[] = [
    { type: 'rawXml', xml: '<w:bookmarkStart b:id="7" w:name="Range"/>' },
    { type: 'paragraph', paraId: '00000002', content: [] },
    { type: 'rawXml', xml: '<w:bookmarkEnd b:id="7"/>' },
  ];
  expect(() => restoreBlockBookmarks([], base, 'cell')).toThrow('Unbound block bookmark attribute b:id');
});
