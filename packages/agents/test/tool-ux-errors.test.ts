import { expect, test } from 'bun:test';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createOfficeMcpServer } from '../src/mcp';
import { openDocx, openPptx } from '../src';
import { fixture } from './fixture';
import { presentationFixture } from './pptx-fixture';

function failure(run: () => unknown): { code?: string; message: string } {
  try { run(); } catch (error) { return error as { code?: string; message: string }; }
  throw new Error('expected a failure');
}

async function asyncFailure(run: () => Promise<unknown>): Promise<{ code?: string; message: string }> {
  try { await run(); } catch (error) { return error as { code?: string; message: string }; }
  throw new Error('expected a failure');
}

test('literal text as a match ID explains that match takes an ID from office_grep', async () => {
  const doc = await openDocx(await fixture());
  try {
    const literal = failure(() => doc.propose({ author: 'test', edits: [{ match: 'using the equations of (Eugster, 1988)', newText: 'x' }] }));
    expect(literal.code).toBe('UNKNOWN_MATCH');
    expect(literal.message).toContain('office_grep');
    expect(literal.message).toContain('such as m1');
    expect(literal.message).toContain('not the text to replace');
    expect(literal.message).toContain('copy the match field');
    const expired = failure(() => doc.propose({ author: 'test', edits: [{ match: 'm999', newText: 'x' }] }));
    expect(expired.code).toBe('UNKNOWN_MATCH');
    expect(expired.message).toContain('Match expired or belongs to another document');
    const formatted = await asyncFailure(() => doc.edit({ version: doc.list().version, edits: [{ op: 'format_text', match: 'the text itself', bold: true }] }));
    expect(formatted.code).toBe('UNKNOWN_MATCH');
    expect(formatted.message).toContain('such as m1');
    expect(formatted.message).toContain('not the text to replace');
    const commented = await asyncFailure(() => doc.edit({ version: doc.list().version, edits: [{ op: 'add_comment', match: 'the text itself', author: 'test', text: 'Note' }] }));
    expect(commented.code).toBe('UNKNOWN_MATCH');
    expect(commented.message).toContain('not the text to replace');
  } finally { doc.close(); }
});

test('PPTX literal text as a match ID uses the deck match ID form', async () => {
  const deck = await openPptx(await presentationFixture());
  try {
    const literal = failure(() => deck.propose({ author: 'test', edits: [{ match: 'Revenue text itself', newText: 'x' }] }));
    expect(literal.code).toBe('UNKNOWN_MATCH');
    expect(literal.message).toContain('office_grep');
    expect(literal.message).toMatch(/such as \S+:match1/u);
    expect(literal.message).toContain('not the text to replace');
    const expired = failure(() => deck.propose({ author: 'test', edits: [{ match: 'some-deck:match999', newText: 'x' }] }));
    expect(expired.code).toBe('UNKNOWN_MATCH');
    expect(expired.message).toContain('Match expired or belongs to another deck');
  } finally { deck.close(); }
});

test('tracked acceptance defaults its date to now and still rejects an invalid explicit date', async () => {
  const doc = await openDocx(await fixture());
  try {
    const hit = doc.grep({ query: '€4.2 million' }).matches[0];
    const proposal = doc.propose({ author: 'test', edits: [{ match: hit.match, newText: '€5 million' }] });
    const before = Date.now() - 5000;
    await doc.accept(proposal.id, { tracked: true });
    const after = Date.now() + 5000;
    const saved = await JSZip.loadAsync(await doc.export());
    const xml = await saved.file('word/document.xml')!.async('string');
    const dates = [...xml.matchAll(/<w:(?:ins|del)\b[^>]*\bw:date="([^"]+)"/gu)].map(match => Date.parse(match[1]));
    expect(dates.length).toBeGreaterThan(0);
    for (const date of dates) { expect(date).toBeGreaterThanOrEqual(before); expect(date).toBeLessThanOrEqual(after); }
  } finally { doc.close(); }
  const other = await openDocx(await fixture());
  try {
    const hit = other.grep({ query: '€4.2 million' }).matches[0];
    const proposal = other.propose({ author: 'test', edits: [{ match: hit.match, newText: '€5 million' }] });
    const invalid = await asyncFailure(() => other.accept(proposal.id, { tracked: true, date: 'not-a-date' }));
    expect(invalid.code).toBe('INVALID_ARGUMENT');
  } finally { other.close(); }
});

async function softBreakFixture(extraParagraphs = '', secondShape?: string) {
  const zip = await JSZip.loadAsync(await presentationFixture());
  const path = 'ppt/slides/slide1.xml';
  const paragraphs = '<a:p><a:r><a:rPr b="0" sz="1800"/><a:t>(Label)</a:t></a:r><a:br/><a:r><a:rPr b="1" sz="2400"/><a:t>Next text</a:t></a:r></a:p>';
  let slide = (await zip.file(path)!.async('string')).replace(/<a:p>[\s\S]*<\/a:p>/u, paragraphs + extraParagraphs);
  if (secondShape) slide = slide.replace('</p:spTree>', `<p:sp><p:nvSpPr><p:cNvPr id="9" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${secondShape}</p:txBody></p:sp></p:spTree>`);
  zip.file(path, slide);
  return zip.generateAsync({ type: 'uint8array' });
}

function firstShape(deck: Awaited<ReturnType<typeof openPptx>>) {
  const slide = deck.outline().slides[0];
  return { slide: slide.id, shape: slide.shapes[0].id };
}

test('PPTX find_replace targets a soft line break and keeps run formatting', async () => {
  const deck = await openPptx(await softBreakFixture());
  try {
    const target = firstShape(deck);
    const hits = deck.grep({ query: 'Next text' });
    expect(hits.matches[0].context).toContain('(Label)\vNext text');
    expect(hits.matches[0].context).not.toContain('\n');
    const proposal = deck.propose({ author: 'test', edits: [{ match: hits.matches[0].match, newText: 'Next text' }] });
    expect(proposal.changes[0].oldText).toBe('Next text');
    const edited = deck.edit({ version: deck.overview().version, edits: [{ op: 'find_replace', shape: target.shape, query: '(Label)\v', replacement: '(Label) \v' }] });
    expect(edited.results[0]).toMatchObject({ changed: true, replacements: 1 });
    const story = deck.readSlide(target).shapes[0].stories[0];
    expect(story.text).toBe('(Label) \vNext text');
    expect(story.paragraphs[0].runs.find(run => run.text === 'Next text')?.formatting.bold).toBe(true);
    expect(story.paragraphs[0].runs[0].formatting.bold).toBe(false);
    const saved = await JSZip.loadAsync(await deck.export());
    const xml = await saved.file('ppt/slides/slide1.xml')!.async('string');
    expect(xml.match(/<a:br\b/gu)).toHaveLength(1);
    expect([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/gu)].map(match => match[1]).join('')).toBe('(Label) Next text');
    expect(xml.indexOf('(Label)')).toBeLessThan(xml.indexOf('<a:br'));
    expect(xml.indexOf('<a:br')).toBeLessThan(xml.indexOf('Next text'));
    expect(xml).toContain('b="0"');
    expect(xml).toContain('b="1"');
  } finally { deck.close(); }
});

test('PPTX find_replace can remove a soft line break and still rejects a paragraph break', async () => {
  const deck = await openPptx(await softBreakFixture());
  try {
    const target = firstShape(deck);
    const rejected = failure(() => deck.edit({ version: deck.overview().version, edits: [{ op: 'find_replace', shape: target.shape, query: '(Label)\n', replacement: 'x' }] }));
    expect(rejected.code).toBe('INVALID_ARGUMENT');
    expect(rejected.message).toContain('use \\u000b for a soft line break');
    const edited = deck.edit({ version: deck.overview().version, edits: [{ op: 'find_replace', shape: target.shape, query: '\v', replacement: ' ' }] });
    expect(edited.results[0]).toMatchObject({ changed: true, replacements: 1 });
    expect(deck.readSlide(target).shapes[0].stories[0].text).toBe('(Label) Next text');
    const saved = await JSZip.loadAsync(await deck.export());
    const xml = await saved.file('ppt/slides/slide1.xml')!.async('string');
    expect(xml).not.toContain('<a:br');
    expect(xml).toContain('Next text');
  } finally { deck.close(); }
});

for (const [name, extraParagraphs] of [
  ['more than 100 paragraphs', Array.from({ length: 100 }, (_, index) => `<a:p><a:pPr algn="ctr" lvl="1" marL="342900"><a:buChar char="*"/></a:pPr><a:r><a:rPr i="1" lang="de-DE"/><a:t>Keep ${index} 😀</a:t></a:r><a:endParaRPr sz="1800"/></a:p>`).join('')],
  ['more than 16000 UTF-16 units', `<a:p><a:pPr algn="r"/><a:r><a:rPr i="1"/><a:t>${'😀'.repeat(8001)}</a:t></a:r></a:p>`],
  ['more than 64 runs in an untouched paragraph', `<a:p><a:pPr algn="ctr"/><a:r><a:t>Keep </a:t></a:r>${Array.from({ length: 65 }, (_, index) => `<a:r><a:rPr b="${index % 2}"/><a:t>${index}😀</a:t></a:r>`).join('')}</a:p>`],
]) test(`PPTX soft-break find_replace preserves unrelated text with ${name}`, async () => {
  const bytes = await softBreakFixture(extraParagraphs);
  const deck = await openPptx(bytes);
  let saved: Uint8Array;
  try {
    const target = firstShape(deck);
    const edited = deck.edit({ version: deck.overview().version, edits: [{ op: 'find_replace', shape: target.shape, query: '(Label)\v', replacement: '(Label) \v' }] });
    expect(edited.results[0]).toMatchObject({ changed: true, replacements: 1, changedShapes: 1 });
    const story = deck.readSlide(target).shapes[0].stories[0];
    expect(story.text).toStartWith('(Label) \vNext text\n');
    expect(story.paragraphs[0].runs[0].formatting.bold).toBe(false);
    expect(story.paragraphs[0].runs.find(run => run.text === 'Next text')?.formatting.bold).toBe(true);
    saved = await deck.export();
    const source = await JSZip.loadAsync(bytes);
    const zip = await JSZip.loadAsync(saved);
    const original = (await source.file('ppt/slides/slide1.xml')!.async('string')).match(/<a:p>[\s\S]*?<\/a:p>/gu)!;
    const paragraphs = (await zip.file('ppt/slides/slide1.xml')!.async('string')).match(/<a:p>[\s\S]*?<\/a:p>/gu)!;
    expect(paragraphs).toHaveLength(original.length);
    expect(paragraphs.slice(1)).toEqual(original.slice(1));
  } finally { deck.close(); }
  const reopened = await openPptx(saved!);
  try { expect(reopened.readSlide(firstShape(reopened)).shapes[0].stories[0].text).toStartWith('(Label) \vNext text\n'); }
  finally { reopened.close(); }
});

test('PPTX find_replace inserts a soft break in a paragraph with more than 64 runs', async () => {
  const zip = await JSZip.loadAsync(await softBreakFixture());
  const path = 'ppt/slides/slide1.xml';
  const runs = Array.from({ length: 65 }, (_, index) => `<a:r><a:rPr i="${index % 2}"/><a:t>${index}😀</a:t></a:r>`).join('');
  zip.file(path, (await zip.file(path)!.async('string')).replace('</a:p>', `${runs}</a:p>`));
  const deck = await openPptx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const target = firstShape(deck);
    const edited = deck.edit({ version: deck.overview().version, edits: [{ op: 'find_replace', shape: target.shape, query: 'Next text', replacement: 'Next\vtext' }] });
    expect(edited.results[0]).toMatchObject({ changed: true, replacements: 1 });
    const story = deck.readSlide(target).shapes[0].stories[0];
    expect(story.text).toBe(`(Label)\vNext\vtext${Array.from({ length: 65 }, (_, index) => `${index}😀`).join('')}`);
    expect(story.paragraphs[0].runs.find(run => run.text === 'Next\vtext')?.formatting.bold).toBe(true);
    const formatted = Array.from({ length: Math.ceil(story.paragraphs[0].runCount / 8) }, (_, page) => deck.readSlide({ ...target, runOffset: page * 8 }).shapes[0].stories[0].paragraphs[0].runs).flat();
    expect(formatted.filter(run => /^\d+😀$/u.test(run.text)).map(run => ({ text: run.text, italic: run.formatting.italic }))).toEqual(Array.from({ length: 65 }, (_, index) => ({ text: `${index}😀`, italic: Boolean(index % 2) })));
    const saved = await JSZip.loadAsync(await deck.export());
    const xml = await saved.file(path)!.async('string');
    expect(xml.match(/<a:br\b/gu)).toHaveLength(2);
  } finally { deck.close(); }
});

test('PPTX soft-break find_replace updates multiple shapes and matches without shifting offsets', async () => {
  const paragraph = '<a:p><a:pPr algn="ctr" lvl="1"/><a:r><a:rPr i="1"/><a:t>😀(Label)</a:t></a:r><a:br/><a:r><a:rPr b="1"/><a:t>Next text (Label)</a:t></a:r><a:br/><a:r><a:rPr i="1"/><a:t>Tail</a:t></a:r></a:p>';
  const deck = await openPptx(await softBreakFixture('', paragraph));
  try {
    const edited = deck.edit({ version: deck.overview().version, edits: [{ op: 'find_replace', query: '(Label)\v', replacement: '(Label) \v' }] });
    expect(edited.results[0]).toMatchObject({ changed: true, replacements: 3, changedShapes: 2 });
    const shapes = deck.readSlide({ slide: firstShape(deck).slide }).shapes;
    expect(shapes[0].stories[0].text).toBe('(Label) \vNext text');
    const story = shapes[1].stories[0];
    expect(story.text).toBe('😀(Label) \vNext text (Label) \vTail');
    expect(story.paragraphs[0]).toMatchObject({ alignment: 'ctr', level: 1 });
    expect(story.paragraphs[0].runs[0].formatting.italic).toBe(true);
    expect(story.paragraphs[0].runs.find(run => run.text === 'Tail')?.formatting.italic).toBe(true);
  } finally { deck.close(); }
});

for (const [name, paragraph, code] of [
  ['a later match contains a field', '<a:p><a:r><a:t>(Label)</a:t></a:r><a:br/><a:fld id="{AE409839-2DAE-4A29-9CE6-CAEC912BBF6D}" type="slidenum"><a:t>1</a:t></a:fld></a:p>', 'UNSUPPORTED_TARGET'],
  ['a later match has unwritable formatting', '<a:p><a:r><a:rPr sz="50"/><a:t>(Label)</a:t></a:r><a:br/><a:r><a:t>Next text</a:t></a:r></a:p><a:p><a:r><a:t>(Label)</a:t></a:r><a:br/><a:r><a:t>Valid tail</a:t></a:r></a:p>', 'INVALID_EDIT'],
]) test(`PPTX soft-break find_replace stays atomic when ${name}`, async () => {
  const deck = await openPptx(await softBreakFixture('', paragraph));
  try {
    const version = deck.overview().version;
    const before = await deck.export();
    const rejected = failure(() => deck.edit({ version, edits: [{ op: 'set_notes', slide: firstShape(deck).slide, text: 'Roll back' }, { op: 'find_replace', query: '(Label)\v', replacement: '(Label) \v' }] }));
    expect(rejected.code).toBe(code);
    expect(rejected.message).toContain('No batch edits were applied.');
    expect(deck.overview().version).toBe(version);
    expect(await deck.export()).toEqual(before);
    expect(deck.readSlide(firstShape(deck)).shapes[0].stories[0].text).toBe('(Label)\vNext text');
  } finally { deck.close(); }
});

test('MCP descriptions name match IDs and the default acceptance date', async () => {
  const root = await mkdtemp(join(tmpdir(), 'office-tool-ux-'));
  await writeFile(join(root, 'report.docx'), await fixture());
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'tool-ux-test', version: '1' });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(b);
  await client.connect(a);
  try {
    const tools = (await client.listTools()).tools;
    const propose = tools.find(tool => tool.name === 'office_propose')!;
    expect((propose.inputSchema.properties!.edits as any).items.properties.match.description).toBe('Match ID from office_grep (not the text).');
    expect(tools.find(tool => tool.name === 'office_accept')!.description).toContain('defaults to now');
    expect(tools.find(tool => tool.name === 'pptx_edit')!.description).toContain('find_replace query and replacement may contain \\u000b for a soft line break');
  } finally { await client.close(); await server.close(); await rm(root, { recursive: true, force: true }); }
});
