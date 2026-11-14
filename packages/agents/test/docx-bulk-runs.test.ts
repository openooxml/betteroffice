import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { attachDocx, openDocx, type DocxAgentDocument, type DocxAgentEdit } from '../src';
import { createOfficeMcpServer } from '../src/mcp';
import { PrototypeDocument } from '../src/prototype';
import { cliEnv } from './cli-env';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rich = '<w:p w:rsidR="00000001"><w:pPr/><w:r w:rsidR="00000002"><w:rPr><w:b/><w:lang w:val="en-US"/></w:rPr><w:t>Alpha</w:t></w:r><w:r w:rsidRPr="00000003"><w:rPr><w:i/><w:color w:val="0066AA"/></w:rPr><w:t>Beta</w:t></w:r><w:r w:rsidR="00000004"><w:rPr><w:b/><w:lang w:val="en-US"/></w:rPr><w:t>Gamma</w:t></w:r><w:r><w:t>Plain</w:t></w:r></w:p>';
const fixed = '<w:p><w:r><w:t>Before</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:tab/></w:r><w:hyperlink w:anchor="target"><w:r><w:t>Link</w:t></w:r></w:hyperlink><w:ins w:id="7" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Inserted</w:t></w:r></w:ins><w:del w:id="8" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Deleted</w:delText></w:r></w:del><w:r><w:t>After</w:t></w:r></w:p>';
const kept = '<w:p w:rsidR="12345678"><w:r><w:t>Keep &#x61;ll bytes.</w:t></w:r></w:p>';
const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

async function fixture(body = rich + fixed + table + kept, header = true) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    (header ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : '') + '</Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  if (header) {
    zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Type="${R}/header" Target="header1.xml"/></Relationships>`);
    zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p><w:r w:rsidR="00000009"><w:rPr><w:b/></w:rPr><w:t>Header</w:t></w:r></w:p></w:hdr>`);
  }
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr>${header ? '<w:headerReference w:type="default" r:id="header"/>' : ''}</w:sectPr></w:body></w:document>`);
  zip.file('customXml/kept.xml', '<kept value="unchanged"/>');
  return zip.generateAsync({ type: 'uint8array' });
}

const documents: DocxAgentDocument[] = [];
async function open(bytes = fixture(), maxEdits?: number) {
  const doc = await openDocx(await bytes, { maxEdits });
  documents.push(doc);
  return doc;
}
afterEach(() => { documents.splice(0).forEach(doc => doc.close()); });

async function xml(bytes: Uint8Array, part = 'word/document.xml') {
  return (await JSZip.loadAsync(bytes)).file(part)!.async('string');
}

const paragraphsOf = (value: string) => value.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
const propertiesOf = (value: string) => (value.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []).map(run => run.match(/<w:rPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:rPr>)/)?.[0] ?? '');

function formattedCharacters(doc: DocxAgentDocument, ref: string, transform = (text: string) => text) {
  const paragraph = doc.read(ref, { length: 16000 });
  return paragraph.runs.flatMap(run => Array.from(transform(paragraph.text.slice(run.start, run.end)), text => ({ text, formatting: run.formatting })));
}

function formattedParagraphs(doc: DocxAgentDocument, transform = (text: string) => text) {
  const paragraphs: Array<ReturnType<typeof formattedCharacters>> = [];
  let offset: number | null = 0;
  do {
    const page = doc.readRuns({ offset, limit: 200 });
    paragraphs.push(...page.paragraphs.map(p => formattedCharacters(doc, p.ref, transform)));
    offset = page.nextOffset;
  } while (offset !== null);
  return paragraphs;
}

function setRuns(paragraph: ReturnType<DocxAgentDocument['readRuns']>['paragraphs'][number], text: (text: string) => string = text => text.toUpperCase()): DocxAgentEdit {
  return { op: 'set_runs', paragraph: paragraph.ref, revision: paragraph.revision,
    runs: paragraph.runs.flatMap(run => 'text' in run ? [{ run: run.id, text: text(run.text) }] : []) };
}

test('readRuns paginates paragraphs, respects boundaries and deduplicates visible formats', async () => {
  const doc = await open();
  const first = doc.readRuns({ limit: 1 });
  expect(first).toMatchObject({ version: doc.list().version, nextOffset: 1, total: 5 });
  const paragraph = first.paragraphs[0];
  const read = doc.read(paragraph.ref);
  expect(paragraph).toMatchObject({ ref: read.ref, revision: read.revision, text: read.text });
  expect(paragraph.runs).toEqual(read.runs.map((run, index) => ({ id: `${read.ref}.${index + 1}`, text: read.text.slice(run.start, run.end),
    ...(Object.keys(run.formatting).length ? { f: index === 1 ? 'f2' : 'f1' } : {}) })));
  expect(first.formats).toEqual({ f1: read.runs[0].formatting, f2: read.runs[1].formatting });
  const bounded = doc.readRuns({ maxChars: paragraph.text.length });
  expect(bounded.paragraphs).toHaveLength(1);
  expect(bounded.nextOffset).toBe(1);
  expect(doc.readRuns({ maxChars: 1 }).paragraphs).toEqual(first.paragraphs);
  const rest = doc.readRuns({ offset: 1, limit: 200 });
  expect(rest.paragraphs).toHaveLength(4);
  expect(rest.nextOffset).toBeNull();
  expect(doc.readRuns({ offset: 5 })).toMatchObject({ paragraphs: [], formats: {}, nextOffset: null, total: 5 });
  const header = doc.list({ limit: 100 }).items.find(p => p.text === 'Header')!;
  expect(doc.readRuns({ story: header.story }).paragraphs.map(p => p.text)).toEqual(['Header']);
  expect(Object.keys(doc.readRuns({ story: header.story }).formats)).toEqual(['f1']);
  for (const options of [{ offset: -1 }, { limit: 0 }, { limit: 201 }, { maxChars: 0 }, { story: 'missing' }]) expect(() => doc.readRuns(options)).toThrow();
});

test('readRuns defaults to 50 paragraphs and 32000 characters', async () => {
  const doc = await open(fixture(Array.from({ length: 210 }, () => '<w:p><w:r><w:t>Short</w:t></w:r></w:p>').join(''), false));
  expect(doc.readRuns().paragraphs).toHaveLength(50);
  expect(doc.readRuns({ limit: 200 }).nextOffset).toBe(200);
  const long = await open(fixture(Array.from({ length: 4 }, () => `<w:p><w:r><w:t>${'a'.repeat(11000)}</w:t></w:r></w:p>`).join(''), false));
  expect(long.readRuns().paragraphs).toHaveLength(2);
  expect(long.readRuns().nextOffset).toBe(2);
});

test('fixed fields, tabs, hyperlinks and tracked runs cannot be set, and offsets use accepted text', async () => {
  const doc = await open();
  const page = doc.readRuns();
  const paragraph = page.paragraphs.find(p => p.text.startsWith('Before'))!;
  const tokens = paragraph.runs.filter(run => 'kind' in run);
  expect(tokens.map(run => 'kind' in run && run.kind)).toEqual(expect.arrayContaining(['field', 'tab', 'protected']));
  expect(tokens.filter(run => 'kind' in run && run.kind === 'protected')).toHaveLength(3);
  for (const token of tokens) {
    await expect(doc.edit({ version: page.version, edits: [{ op: 'set_runs', paragraph: paragraph.ref, revision: paragraph.revision, runs: [{ run: token.id, text: 'changed' }] }] })).rejects.toMatchObject({ code: 'PROTECTED_RUN' });
  }
  const last = paragraph.runs.at(-1)!;
  await doc.edit({ version: page.version, edits: [{ op: 'set_runs', paragraph: paragraph.ref, revision: paragraph.revision, runs: [{ run: last.id, text: 'Later' }] }] });
  expect(await xml(await doc.export())).toContain(fixed.replace('>After<', '>Later<'));
});

test('a tab inside a text run remains fixed while both surrounding spans can change', async () => {
  const content = '<w:p><w:r w:rsidR="00000002"><w:rPr><w:b/></w:rPr><w:t>Left</w:t><w:tab/><w:t>Right</w:t></w:r></w:p>';
  const bytes = await fixture(content, false);
  const doc = await open(Promise.resolve(bytes));
  const page = doc.readRuns();
  const p = page.paragraphs[0];
  expect(p.runs).toEqual([{ id: `${p.ref}.1`, text: 'Left', f: 'f1' }, { id: `${p.ref}.2`, kind: 'tab' }, { id: `${p.ref}.3`, text: 'Right', f: 'f1' }]);
  await doc.edit({ version: page.version, edits: [setRuns(p)] });
  expect(await xml(await doc.export())).toBe((await xml(bytes)).replace('>Left<', '>LEFT<').replace('>Right<', '>RIGHT<'));
});

test('set_runs retains rPr, rsids and paragraph bytes for a single run edit', async () => {
  const bytes = await fixture();
  const doc = await open(Promise.resolve(bytes));
  const page = doc.readRuns();
  const p = page.paragraphs[0];
  await doc.edit({ version: page.version, edits: [{ op: 'set_runs', paragraph: p.ref, revision: p.revision, runs: [
    { run: p.runs[1].id, text: 'BETA' },
  ] }] });
  expect(await xml(await doc.export())).toBe((await xml(bytes)).replace('>Beta<', '>BETA<'));
});

test('set_runs matches proposal saves for rewritten paragraphs and preserves character formatting', async () => {
  const bytes = await fixture();
  const source = await xml(bytes);
  const doc = await open(Promise.resolve(bytes));
  const control = await open(Promise.resolve(bytes));
  const page = doc.readRuns();
  const p = page.paragraphs[0];
  const texts = ['A longer alpha', 'B', 'Γάμμα 😀'];
  const original = doc.read(p.ref);
  const expected = original.runs.flatMap((run, index) => Array.from(texts[index] ?? original.text.slice(run.start, run.end), text => ({ text, formatting: run.formatting })));
  await doc.edit({ version: page.version, edits: [{ op: 'set_runs', paragraph: p.ref, revision: p.revision, runs: [
    { run: p.runs[0].id, text: texts[0] }, { run: p.runs[1].id, text: texts[1] }, { run: p.runs[2].id, text: texts[2] },
  ] }] });
  for (let index = 2; index >= 0; index--) {
    const run = original.runs[index];
    const query = original.text.slice(run.start, run.end);
    const match = control.grep({ query, caseSensitive: true }).matches[0];
    const proposal = control.propose({ author: 'Test', edits: [{ match: match.match, newText: texts[index] }] });
    await control.accept(proposal.id);
  }
  const saved = await doc.export();
  const savedXml = await xml(saved);
  const controlXml = await xml(await control.export());
  expect(savedXml).toBe(controlXml);
  expect(propertiesOf(savedXml)).toEqual(propertiesOf(controlXml));
  expect(formattedCharacters(doc, p.ref)).toEqual(expected);
  const reopened = await open(Promise.resolve(saved));
  expect(formattedCharacters(reopened, reopened.readRuns().paragraphs[0].ref)).toEqual(expected);
  expect(paragraphsOf(savedXml).slice(1)).toEqual(paragraphsOf(source).slice(1));
  expect(await xml(saved, 'word/header1.xml')).toBe(await xml(bytes, 'word/header1.xml'));
  expect(await xml(saved, 'customXml/kept.xml')).toBe(await xml(bytes, 'customXml/kept.xml'));
  const fresh = doc.readRuns().paragraphs[0];
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_runs', paragraph: fresh.ref, revision: fresh.revision, runs: [{ run: fresh.runs[1].id, text: '' }] }] });
  expect(doc.read(fresh.ref).text).toBe('A longer alphaΓάμμα 😀Plain');
  const deleted = await xml(await doc.export());
  expect(deleted).not.toContain('>B<');
  expect(deleted).toContain(kept);
});

test('set_runs validates revisions, ids, duplicate targets and text atomically', async () => {
  const doc = await open();
  const page = doc.readRuns();
  const p = page.paragraphs[0];
  const run = p.runs[0].id;
  const edit = (runs: Array<{ run: string; text: string }>, revision = p.revision): DocxAgentEdit => ({ op: 'set_runs', paragraph: p.ref, revision, runs });
  for (const [invalid, code] of [
    [edit([{ run, text: 'new' }], 0), 'STALE_TARGET'],
    [edit([{ run, text: 'new' }], p.revision + 1), 'STALE_TARGET'],
    [edit([{ run: `${p.ref}.999`, text: 'new' }]), 'UNKNOWN_REF'],
    [edit([{ run: `${p.ref}.01`, text: 'new' }]), 'UNKNOWN_REF'],
    [edit([{ run: page.paragraphs[1].runs[0].id, text: 'new' }]), 'UNKNOWN_REF'],
    [edit([{ run, text: 'one' }, { run, text: 'two' }]), 'INVALID_ARGUMENT'],
    ...['\n', '\r', '\t', '\ufffc', '\ud800'].map(text => [edit([{ run, text }]), 'INVALID_ARGUMENT'] as const),
  ] as const) {
    const before = await xml(await doc.export());
    await expect(doc.edit({ version: page.version, edits: [setRuns(page.paragraphs[2]), invalid] })).rejects.toMatchObject({ code });
    expect(await xml(await doc.export())).toBe(before);
    expect(doc.list().version).toBe(page.version);
  }
  await expect(doc.edit({ version: page.version, edits: [edit([{ run, text: 'one' }]), edit([{ run, text: 'two' }])] })).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  const receipt = await doc.edit({ version: page.version, edits: [edit([{ run, text: 'changed' }])] });
  await expect(doc.edit({ version: receipt.version, edits: [edit([{ run, text: 'stale' }])] })).rejects.toMatchObject({ code: 'STALE_TARGET' });
});

test('set_runs rolls back if an earlier step removed its target and supports table and header stories', async () => {
  const bytes = await fixture();
  const doc = await open(Promise.resolve(bytes));
  const page = doc.readRuns();
  await expect(doc.edit({ version: page.version, edits: [{ op: 'delete_paragraph', paragraph: page.paragraphs[0].ref }, setRuns(page.paragraphs[0])] })).rejects.toMatchObject({ code: 'UNKNOWN_REF' });
  expect(await xml(await doc.export())).toBe(await xml(bytes));
  const cell = page.paragraphs.find(p => p.text === 'Cell')!;
  const header = page.paragraphs.find(p => p.text === 'Header')!;
  await doc.edit({ version: page.version, edits: [setRuns(cell), setRuns(header)] });
  expect(await xml(await doc.export())).toBe((await xml(bytes)).replace('>Cell<', '>CELL<'));
  expect(await xml(await doc.export(), 'word/header1.xml')).toBe((await xml(bytes, 'word/header1.xml')).replace('>Header<', '>HEADER<'));
});

test('run ids go stale when an accepted proposal changes the paragraph, even with its revision preserved', async () => {
  const owner = await open();
  const agent = attachDocx(owner.session);
  try {
    const page = agent.readRuns();
    const p = page.paragraphs[0];
    const proposal = agent.propose({ author: 'agent', edits: [{ match: agent.grep({ query: 'Alpha' }).matches[0].match, newText: '' }] });
    await agent.accept(proposal.id);
    await expect(agent.edit({ version: agent.list().version, edits: [{ op: 'set_runs', paragraph: p.ref, revision: p.revision, runs: [{ run: p.runs[0].id, text: 'Wrong' }] }] }))
      .rejects.toMatchObject({ code: 'STALE_TARGET' });
    const fresh = agent.readRuns().paragraphs[0];
    expect(fresh.revision).not.toBe(p.revision);
    await agent.edit({ version: agent.list().version, edits: [{ op: 'set_runs', paragraph: fresh.ref, revision: fresh.revision, runs: [{ run: fresh.runs[0].id, text: 'Right' }] }] });
    expect(agent.readRuns().paragraphs[0].text.startsWith('Right')).toBe(true);
  } finally { agent.close(); }
});

test('set_runs refuses a batch with an offset based edit in the same paragraph', async () => {
  const doc = await open();
  const page = doc.readRuns();
  const p = page.paragraphs[0];
  const match = doc.grep({ query: 'Beta' }).matches[0].match;
  const source = await xml(await doc.export());
  await expect(doc.edit({ version: page.version, edits: [
    { op: 'set_runs', paragraph: p.ref, revision: p.revision, runs: [{ run: p.runs[0].id, text: 'Much longer alpha' }] },
    { op: 'format_text', match, bold: true },
  ] })).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  expect(await xml(await doc.export())).toBe(source);
  const other = page.paragraphs.find(paragraph => paragraph.text === 'Cell')!;
  const applied = await doc.edit({ version: page.version, edits: [
    { op: 'set_runs', paragraph: other.ref, revision: other.revision, runs: [{ run: other.runs[0].id, text: 'Zelle' }] },
    { op: 'format_text', match, bold: true },
  ] });
  expect(applied.applied).toBe(true);
});

test('100 sequential proposals succeed while the 65th pending proposal refuses', async () => {
  const doc = await open(fixture('<w:p><w:r><w:t>word</w:t></w:r></w:p>', false));
  for (let i = 0; i < 100; i++) {
    const match = doc.grep({ query: i % 2 ? 'WORD' : 'word', caseSensitive: true }).matches[0];
    const proposal = doc.propose({ author: 'Test', edits: [{ match: match.match, newText: i % 2 ? 'word' : 'WORD' }] });
    await doc.accept(proposal.id);
  }
  expect(doc.readRuns().paragraphs[0].text).toBe('word');
  const match = doc.grep({ query: 'word' }).matches[0];
  for (let i = 0; i < 64; i++) doc.propose({ author: 'Test', edits: [{ match: match.match, newText: 'pending' }] });
  try { doc.propose({ author: 'Test', edits: [{ match: match.match, newText: 'extra' }] }); throw new Error('Expected proposal limit.'); }
  catch (error) { expect(error).toMatchObject({ code: 'PROPOSAL_LIMIT' }); }
});

class ProposalProbe extends PrototypeDocument<string, string> {
  readonly format = 'pptx' as const;
  readonly name = 'probe';
  constructor(maxEdits = 32) { super(maxEdits); }
  protected currentVersion() { return '1'; }
  protected validate() {}
  protected apply() {}
  protected save() { return new Uint8Array(); }
  protected reopen() {}
  protected dispose() {}
  propose(changes: string[]) { return this.stage({ author: 'Test' }, changes, changes, '1'); }
}

test('prototype proposal count and character caps scale with maxEdits', async () => {
  expect(() => new ProposalProbe().propose(Array(33).fill('word'))).toThrow();
  expect(() => new ProposalProbe().propose(['a'.repeat(32001)])).toThrow();
  expect(new ProposalProbe(200).propose(Array(200).fill('a'.repeat(500))).changes).toHaveLength(200);
  expect(() => new ProposalProbe(200).propose(['a'.repeat(200001)])).toThrow();
  const probe = new ProposalProbe();
  for (let i = 0; i < 100; i++) await probe.accept(probe.propose(['word']).id);
  for (let i = 0; i < 64; i++) probe.propose(['word']);
  try { probe.propose(['word']); throw new Error('Expected proposal limit.'); }
  catch (error) { expect(error).toMatchObject({ code: 'PROPOSAL_LIMIT' }); }
});

async function connect(bytes: Uint8Array, options: { maxEdits?: number; readOnly?: boolean } = {}, cliArgs?: string[]) {
  const root = await mkdtemp(join(tmpdir(), 'docx-bulk-'));
  await writeFile(join(root, 'source.docx'), bytes);
  const client = new Client({ name: 'bulk-runs-test', version: '1' });
  const server = cliArgs ? undefined : await createOfficeMcpServer({ root, ...options });
  if (server) {
    const [a, b] = InMemoryTransport.createLinkedPair();
    await server.connect(b);
    await client.connect(a);
  } else {
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root, ...cliArgs!], env: cliEnv(), stderr: 'pipe' }));
  }
  let calls = 0;
  const call = async (name: string, args: Record<string, unknown>) => {
    calls++;
    const output = await client.callTool({ name, arguments: args });
    expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
    return output.structuredContent as Record<string, any>;
  };
  return { client, call, root, calls: () => calls, close: async () => { await client.close(); await server?.close(); await rm(root, { recursive: true, force: true }); } };
}

test('MCP lists docx_read_runs in read-only mode and validates pagination', async () => {
  const c = await connect(await fixture(), { readOnly: true });
  try {
    const tools = (await c.client.listTools()).tools;
    expect(tools.find(tool => tool.name === 'docx_read_runs')?.annotations?.readOnlyHint).toBe(true);
    expect(tools.map(tool => tool.name)).not.toContain('docx_edit');
    const { document } = await c.call('office_open', { path: 'source.docx' });
    expect(await c.call('docx_read_runs', { document, limit: 1 })).toMatchObject({ nextOffset: 1, total: 5 });
  } finally { await c.close(); }
});

test.each([undefined, 200, 'cli'] as const)('MCP enforces the configured edit cap: %s', async setting => {
  const body = Array.from({ length: 33 }, (_, i) => `<w:p><w:r><w:t>Word${i}</w:t></w:r></w:p>`).join('');
  const c = await connect(await fixture(body, false), { maxEdits: typeof setting === 'number' ? setting : undefined }, setting === 'cli' ? ['--max-edits', '200'] : undefined);
  try {
    const { document } = await c.call('office_open', { path: 'source.docx' });
    const page = await c.call('docx_read_runs', { document });
    const edits = page.paragraphs.map((p: Parameters<typeof setRuns>[0]) => setRuns(p));
    const matches = (await c.call('office_grep', { document, query: 'Word', limit: 100 })).matches;
    const proposed = await c.client.callTool({ name: 'office_propose', arguments: { document, author: 'Test', edits: matches.map((m: { match: string }) => ({ match: m.match, newText: 'WORD' })) } });
    const direct = await c.client.callTool({ name: 'docx_edit', arguments: { document, version: page.version, edits } });
    if (setting === undefined) {
      expect(proposed.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT' });
      expect(direct.structuredContent).toMatchObject({ code: 'INVALID_ARGUMENT' });
    } else {
      expect(proposed.isError, JSON.stringify(proposed.structuredContent)).not.toBe(true);
      expect(direct.isError, JSON.stringify(direct.structuredContent)).not.toBe(true);
    }
    const tools = (await c.client.listTools()).tools;
    for (const name of ['docx_edit', 'office_propose']) expect(tools.find(tool => tool.name === name)?.description).toContain(String(setting === undefined ? 32 : 200));
  } finally { await c.close(); }
}, 30000);

test('MCP maxEdits rejects values outside 1 to 500', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-bulk-cap-'));
  try {
    for (const maxEdits of [0, -1, 501, 1.5, NaN, Infinity]) await expect(createOfficeMcpServer({ root, maxEdits })).rejects.toThrow();
  } finally { await rm(root, { recursive: true, force: true }); }
});

test('310 paragraph translation preserves text and formatting with fewer real MCP calls', async () => {
  const body = Array.from({ length: 309 }, (_, i) => `${i === 308 ? table : ''}<w:p w:rsidR="00000001"><w:r w:rsidR="00000002"><w:rPr><w:b/></w:rPr><w:t>alpha${i}</w:t></w:r><w:r w:rsidRPr="00000003"><w:rPr><w:i/><w:color w:val="0066AA"/></w:rPr><w:t>beta${i}</w:t></w:r><w:r><w:rPr><w:color w:val="AA6600"/></w:rPr><w:t>gamma${i}</w:t></w:r>${i % 31 === 0 ? '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' : ''}</w:p>`).join('');
  const bytes = await fixture(body, false);
  const before = await connect(bytes);
  const after = await connect(bytes, { maxEdits: 200 });
  try {
    const old = await before.call('office_open', { path: 'source.docx' });
    const replacements = ['alpha', 'beta', 'gamma'].flatMap(word => Array.from({ length: 309 }, (_, i) => ({ query: `${word}${i}`, newText: `${word}${i}`.toUpperCase() })));
    replacements.push({ query: 'Cell', newText: 'CELL' });
    const batches: Array<typeof replacements> = [];
    let batch: typeof replacements = [];
    let chars = 0;
    for (const replacement of replacements) {
      const size = replacement.query.length + replacement.newText.length;
      if (batch.length === 32 || chars + size > 32000) {
        batches.push(batch);
        batch = [];
        chars = 0;
      }
      batch.push(replacement);
      chars += size;
    }
    if (batch.length) batches.push(batch);
    expect(replacements).toHaveLength(928);
    expect(batches.map(batch => batch.length)).toEqual(Array(29).fill(32));
    for (const batch of batches) {
      const edits: Array<{ match: string; newText: string }> = [];
      for (const { query, newText } of batch) {
        const found = await before.call('office_grep', { document: old.document, query, caseSensitive: true, limit: 100 });
        const match = found.matches.find((m: { text: string; context: string; start: number; contextStart: number }) => !/\d/.test(m.context[m.start - m.contextStart + query.length] ?? ''));
        expect(match).toBeDefined();
        edits.push({ match: match.match, newText });
      }
      const proposal = await before.call('office_propose', { document: old.document, author: 'Translator', edits });
      await before.call('office_review', { document: old.document, proposal: proposal.id });
      await before.call('office_accept', { document: old.document, proposal: proposal.id });
    }
    const fresh = await after.call('office_open', { path: 'source.docx' });
    let offset: number | null = 0;
    let version = fresh.version;
    do {
      const page = await after.call('docx_read_runs', { document: fresh.document, offset, limit: 200 });
      expect(page.total).toBe(310);
      const edited = await after.call('docx_edit', { document: fresh.document, version, edits: page.paragraphs.map((p: Parameters<typeof setRuns>[0]) => setRuns(p)) });
      version = edited.version;
      offset = page.nextOffset;
    } while (offset !== null);
    await before.call('office_export', { document: old.document, path: 'translated.docx' });
    await after.call('office_export', { document: fresh.document, path: 'translated.docx' });
    const oldCalls = before.calls(), newCalls = after.calls();
    expect(oldCalls).toBe(2 + replacements.length + batches.length * 3);
    expect(newCalls).toBe(2 + Math.ceil(310 / 200) * 2);
    expect(newCalls).toBeLessThan(oldCalls / 10);
    console.log(`TOOLCALLS before=${oldCalls} after=${newCalls}`);
    const oldBytes = await readFile(join(before.root, 'translated.docx'));
    const newBytes = await readFile(join(after.root, 'translated.docx'));
    const oldXml = await xml(oldBytes);
    const newXml = await xml(newBytes);
    const expected = (await xml(bytes)).replace(/>(alpha\d+|beta\d+|gamma\d+|Cell)</g, (_, text: string) => `>${text.toUpperCase()}<`);
    expect(newXml).toBe(oldXml);
    expect(propertiesOf(newXml)).toEqual(propertiesOf(expected));
    expect(paragraphsOf(newXml).filter(p => !p.includes('<w:fldSimple'))).toEqual(paragraphsOf(expected).filter(p => !p.includes('<w:fldSimple')));
    const textNodes = (value: string) => [...value.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]);
    expect(textNodes(newXml)).toEqual(textNodes(expected));
    const sourceDoc = await open(Promise.resolve(bytes));
    const savedDoc = await open(Promise.resolve(newBytes));
    expect(formattedParagraphs(savedDoc)).toEqual(formattedParagraphs(sourceDoc, text => text.toUpperCase()));
  } finally { await before.close(); await after.close(); }
}, 300000);
