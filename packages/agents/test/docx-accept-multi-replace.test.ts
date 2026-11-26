import { expect, spyOn, test } from 'bun:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import JSZip from 'jszip';
import { captureSessionSave, writeSessionSave, yrsToDocument } from '@betteroffice/docx/yrs';
import { attachDocx, openDocx, type DocxAgentDocument } from '../src';
import { createOfficeMcpServer } from '../src/mcp';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DOCUMENT = 'word/document.xml';

type Replacement = { query: string; newText: string };
type RunText = [text: string, italic: boolean];
type CheckParagraph = (paragraph: string, bytes: Uint8Array) => void | Promise<void>;

function run(id: number, text: string, italic = false): string {
  const rsid = id.toString(16).padStart(8, '0');
  const properties = italic ? '<w:rPr><w:i/></w:rPr>' : '';
  return `<w:r w:rsidR="${rsid}">${properties}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}

const before = `<w:p w:rsidR="98765432">${run(0xa1, 'Keep &#x61;ll bytes.')}</w:p>`;
const after = `<w:p w:rsidR="12345678">${run(0xa2, 'Keep this paragraph &#x74;oo.')}</w:p>`;

async function fixture(runs: readonly string[], ending = after): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="doc" Type="${R}/officeDocument" Target="${DOCUMENT}"/></Relationships>`);
  zip.file(DOCUMENT, `<w:document xmlns:w="${W}"><w:body>${before}<w:p>${runs.join('')}</w:p>${ending}<w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const RUN = /<w:r\b(?:[^>]*?\/>|[^>]*>[\s\S]*?<\/w:r>)/g;
const runsOf = (xml: string) => xml.match(RUN) ?? [];

function runText(xml: string, tag: 't' | 'delText' = 't'): RunText {
  const text = [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*>([^<]*)</w:${tag}>`, 'g'))]
    .map(match => match[1]).join('');
  const italic = xml.match(/<w:i\b[^>]*>/)?.[0];
  return [text, italic !== undefined && !/\bw:val="(?:0|false|off)"/.test(italic)];
}

function expectRuns(paragraph: string, untouched: readonly string[], expected: RunText[]) {
  const runs = runsOf(paragraph);
  expect(runs.map(xml => runText(xml)), paragraph).toEqual(expected);
  for (const xml of untouched) expect(runs).toContain(xml);
  expect(paragraph).not.toMatch(/<w:(ins|del)\b/);
}

async function documentXml(bytes: Uint8Array) {
  return (await JSZip.loadAsync(bytes)).file(DOCUMENT)!.async('string');
}

async function compareWithBaseExport(doc: DocxAgentDocument, untouched: readonly string[]) {
  const draft = await doc['fork']();
  try {
    const base = draft.materializeDocx()!;
    const capture = captureSessionSave(draft, draft.paragraphIdentities(), false);
    const basePart = capture.plan.splicedParts?.find(part => part.part === DOCUMENT);
    const localPart = draft.paragraphSavePlan().splicedParts?.find(part => part.part === DOCUMENT);
    const basePatched = basePart?.paragraphPatches?.map(patch => patch.ordinal) ?? [];
    const localPatched = localPart?.paragraphPatches?.map(patch => patch.ordinal) ?? [];
    expect(localPatched).toEqual(expect.arrayContaining(basePatched));
    const { bytes: baseBytes } = await writeSessionSave(draft, yrsToDocument(draft, base), capture,
      base.originalBuffer!, { updateModifiedDate: false });
    const bytes = await doc['exportBytes'](undefined, false);
    const xml = await documentXml(bytes);
    for (const sourceRun of untouched) expect(runsOf(xml)).toContain(sourceRun);
    return { xml, bytes, baseBytes, localPatched, basePatched };
  } finally { draft.destroy(); }
}

async function checkExport(
  save: () => Promise<Uint8Array>, check: CheckParagraph, expectedReplaceCounts: string[][] = [],
) {
  const diagnostic = spyOn(console, 'debug').mockImplementation(() => {});
  try {
    const bytes = await save();
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file(DOCUMENT)!.async('string');
    const paragraphs = xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? [];
    expect(paragraphs).toHaveLength(3);
    expect([paragraphs[0], paragraphs[2]]).toEqual([before, after]);
    await check(paragraphs[1]!, bytes);
    expect(diagnostic.mock.calls.filter(([warning]) =>
      typeof warning === 'string' && warning.startsWith('DOCX_MINIMAL_PATCH replace counts')
    )).toEqual(expectedReplaceCounts);
  } finally { diagnostic.mockRestore(); }
}

async function checkProposal(
  runs: readonly string[], replacements: readonly Replacement[], check: CheckParagraph,
  options: { tracked: boolean; date?: string } = { tracked: false },
  expectedReplaceCounts: string[][] = [],
) {
  const doc = await openDocx(await fixture(runs));
  try {
    const edits = replacements.map(({ query, newText }) => {
      const { matches } = doc.grep({ query });
      expect(matches).toHaveLength(1);
      return { match: matches[0]!.match, newText };
    });
    const proposal = doc.propose({ author: 'test', edits });
    expect(await doc.accept(proposal.id, options)).toMatchObject({
      status: 'accepted', tracked: options.tracked, changes: edits.length,
    });
    await checkExport(() => doc.export(), check, expectedReplaceCounts);
  } finally { doc.close(); }
}

const contractorEdits = [
  { query: '(name of the contractor)', newText: 'Great Lakes Analytical Services' },
  { query: '(name city and state)', newText: 'Chicago, Illinois' },
] as const;
const [contractor, city] = contractorEdits;
const contractorRuns = [
  run(1, 'The contractor is '), run(2, contractor.query, true), run(3, ', ', true),
  run(4, 'located'), run(5, ' '), run(6, 'in '), run(7, `${city.query}.`, true),
] as const;
const contractorUntouched = [contractorRuns[0], ...contractorRuns.slice(2, 6)];
const contractorExpected: RunText[] = [
  ['The contractor is ', false], [contractor.newText, true], [', ', true],
  ['located', false], [' ', false], ['in ', false], [`${city.newText}.`, true],
];

test('MCP acceptance preserves untouched runs across two replacements in one paragraph', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-multi-replace-'));
  const client = new Client({ name: 'docx-multi-replace', version: '1' });
  try {
    await writeFile(join(root, 'input.docx'), await fixture(contractorRuns));
    const server = await createOfficeMcpServer({ root });
    try {
      const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
      await server.connect(serverTransport);
      await client.connect(clientTransport);
      const call = async <T,>(name: string, args: Record<string, unknown>) => {
        const output = await client.callTool({ name, arguments: args });
        expect(output.isError, JSON.stringify(output.structuredContent)).not.toBe(true);
        return output.structuredContent as T;
      };
      const { document } = await call<{ document: string }>('office_open', { path: 'input.docx' });
      const edits = [];
      for (const { query, newText } of contractorEdits) {
        const { matches } = await call<{ matches: Array<{ match: string }> }>('office_grep', { document, query });
        expect(matches).toHaveLength(1);
        edits.push({ match: matches[0]!.match, newText });
      }
      const { id: proposal } = await call<{ id: string }>('office_propose', { document, author: 'test', edits });
      expect(await call('office_accept', { document, proposal, tracked: false })).toMatchObject({
        status: 'accepted', tracked: false, changes: 2,
      });
      await checkExport(async () => {
        await call('office_export', { document, path: 'output.docx' });
        return readFile(join(root, 'output.docx'));
      }, paragraph => expectRuns(paragraph, contractorUntouched, contractorExpected));
    } finally {
      await client.close();
      await server.close();
    }
  } finally { await rm(root, { recursive: true, force: true }); }
}, 30_000);

test('document acceptance preserves untouched runs across two replacements in one paragraph', async () => {
  await checkProposal(contractorRuns, contractorEdits, paragraph =>
    expectRuns(paragraph, contractorUntouched, contractorExpected));
}, 30_000);

test('preserves untouched runs between replacements when every run has the same formatting', async () => {
  const runs = [
    run(1, 'The contractor is '), run(2, contractor.query), run(3, ', '),
    run(4, 'located'), run(5, ' '), run(6, 'in '), run(7, `${city.query}.`),
  ];
  await checkProposal(runs, contractorEdits, paragraph =>
    expectRuns(paragraph, [runs[0]!, ...runs.slice(2, 6)], contractorExpected.map(([text]): RunText => [text, false])));
}, 30_000);

test('patches each replaced run when recorded edits resolve ambiguous alignment', async () => {
  const untouched = run(2, 'X', true);
  await checkProposal([run(1, 'ABCD'), untouched, run(3, 'efgh')], [
    { query: 'ABCD', newText: 'e-f-g-h' }, { query: 'efgh', newText: 'ABCD' },
  ], paragraph => expectRuns(paragraph, [untouched], [
    ['e-f-g-h', false], ['X', true], ['ABCD', false],
  ]));
}, 30_000);

test('accepts a lone recorded insertion and keeps the untouched run byte-identical', async () => {
  const untouched = run(2, ' ti').replace('<w:t', '<w:rPr><w:noProof/></w:rPr><w:t');
  await checkProposal([run(1, 'fi'), untouched], [{ query: 'fi', newText: 'fii' }],
    paragraph => expectRuns(paragraph, [untouched], [['fii', false], [' ti', false]]));
}, 30_000);

test('accepting 33 replacements retains every KEEP run byte-identically', async () => {
  const kept = Array.from({ length: 32 }, (_, i) => run(2 * i + 2, ' K&#x45;EP ', true));
  const runs = kept.flatMap((xml, i) => [run(2 * i + 1, 'a'), xml]);
  runs.push(run(65, 'a'));
  const query = Array(33).fill('a').join(' KEEP ');
  const newText = Array(33).fill('b').join(' KEEP ');
  await checkProposal(runs, [{ query, newText }], paragraph => {
    for (const sourceRun of kept) expect(runsOf(paragraph)).toContain(sourceRun);
    expect(runsOf(paragraph).filter(xml => runText(xml)[0] === 'b')).toHaveLength(33);
  });
}, 30_000);

for (const kind of ['headless', 'attached', 'set_runs', 'crossing runs'] as const) {
  test(`${kind} DOCX export keeps every paragraph patched by base alignment`, async () => {
    const left = run(1, 'Before ');
    const right = run(4, ' after.', true);
    const owner = await openDocx(await fixture([left, run(2, 'alpha'), run(3, ' beta', true), right]));
    const doc = kind === 'attached' ? attachDocx(owner.session) : owner;
    try {
      if (kind === 'set_runs') {
        const page = doc.readRuns();
        const p = page.paragraphs[1]!;
        await doc.edit({ version: page.version, edits: [{ op: 'set_runs', paragraph: p.ref, revision: p.revision,
          runs: [{ run: p.runs[1]!.id, text: 'FIRST' }] }] });
      } else {
        const query = kind === 'crossing runs' ? 'alpha beta' : 'alpha';
        const [hit] = doc.grep({ query }).matches;
        const proposal = doc.propose({ author: 'Agent', edits: [{ match: hit!.match, newText: 'FIRST' }] });
        if (kind === 'attached') expect(owner.session.listRevisions().length).toBeGreaterThan(0);
        expect(await doc.accept(proposal.id)).toMatchObject({ status: 'accepted' });
      }
      const untouched = [left, right, ...(kind === 'crossing runs' ? [] : [run(3, ' beta', true)])];
      const { xml, bytes, baseBytes, localPatched, basePatched } = await compareWithBaseExport(doc, untouched);
      expect(localPatched).toContain(1);
      expect(xml).toContain('FIRST');
      if (kind === 'attached') {
        expect(basePatched).toContain(1);
        expect(bytes).toEqual(baseBytes);
      }
    } finally { doc.close(); if (doc !== owner) owner.close(); }
  }, 30_000);
}

test('successive untracked DOCX acceptances preserve untouched runs in both paragraphs', async () => {
  const firstKept = [run(1, 'a', true), run(4, ' ret&#x61;ined first')];
  const secondKept = [run(5, 'b', true), run(8, ' ret&#x61;ined second')];
  const first = [firstKept[0]!, run(2, 'a'), run(3, 'X', true), firstKept[1]!];
  const second = [secondKept[0]!, run(6, 'b'), run(7, 'X', true), secondKept[1]!];
  const doc = await openDocx(await fixture(first, `<w:p>${second.join('')}</w:p>`));
  try {
    for (const [query, newText] of [['aaX', 'aY'], ['bbX', 'bZ']] as const) {
      const [hit] = doc.grep({ query }).matches;
      const proposal = doc.propose({ author: 'Agent', edits: [{ match: hit!.match, newText }] });
      expect(await doc.accept(proposal.id)).toMatchObject({ status: 'accepted', tracked: false });
    }
    const { xml, localPatched } = await compareWithBaseExport(doc, [...firstKept, ...secondKept]);
    expect(localPatched).toEqual([1, 2]);
    expect(runsOf(xml).map(run => runText(run)[0])).toEqual(expect.arrayContaining(['Y', 'Z']));
  } finally { doc.close(); }
}, 30_000);

test('accepts replacements in neighbouring runs one character apart', async () => {
  const left = run(1, 'Before ');
  const gap = run(3, '-', true);
  const right = run(5, ' after.');
  await checkProposal([left, run(2, 'alpha', true), gap, run(4, 'beta'), right], [
    { query: 'alpha', newText: 'FIRST' }, { query: 'beta', newText: 'SECOND' },
  ], paragraph => expectRuns(paragraph, [left, gap, right], [
    ['Before ', false], ['FIRST', true], ['-', true], ['SECOND', false], [' after.', false],
  ]));
}, 30_000);

test('edits partially replaced runs in place at both paragraph ends', async () => {
  const middle = run(2, 'and');
  await checkProposal([run(1, 'Start first ', true), middle, run(3, ' last finish', true)], [
    { query: 'Start', newText: 'Restart' }, { query: 'finish', newText: 'fiend' },
  ], paragraph => expectRuns(paragraph, [middle], [
    ['Restart first ', true], ['and', false], [' last fiend', true],
  ]));
}, 30_000);

test('tracked acceptance rewrites the edited paragraph and keeps its neighbours', async () => {
  const doc = await openDocx(await fixture(contractorRuns));
  try {
    const edits = contractorEdits.map(({ query, newText }) => ({ match: doc.grep({ query }).matches[0]!.match, newText }));
    const proposal = doc.propose({ author: 'Agent', edits });
    await doc.accept(proposal.id, { tracked: true, date: '2026-01-01T00:00:00Z' });
    const part = doc.session.paragraphSavePlan().splicedParts?.find(part => part.part === DOCUMENT);
    expect(part?.changed).toEqual([1]);
    expect(part?.paragraphPatches ?? []).toEqual([]);
    expect(part?.legacyParagraphs?.map(paragraph => paragraph.ordinal)).toEqual([1]);
    const xml = await documentXml(await doc.export());
    const paragraphs = xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? [];
    expect([paragraphs[0], paragraphs[2]]).toEqual([before, after]);
    expect(paragraphs[1]?.match(/<w:(ins|del)\b/g)).toHaveLength(4);
  } finally { doc.close(); }
}, 30_000);

test('an insertion after a leading inline drawing rewrites only that paragraph', async () => {
  const drawing = '<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:rect style="width:20pt;height:10pt"/></w:pict></w:r>';
  const doc = await openDocx(await fixture([drawing, run(1, 'tail')]));
  try {
    const [hit] = doc.grep({ query: 'tail' }).matches;
    const proposal = doc.propose({ author: 'Agent', edits: [{ match: hit!.match, newText: 'First tail' }] });
    await doc.accept(proposal.id);
    const part = doc.session.paragraphSavePlan().splicedParts?.find(part => part.part === DOCUMENT);
    expect(part?.changed).toEqual([1]);
    expect(part?.paragraphPatches ?? []).toEqual([]);
    expect(part?.legacyParagraphs).toMatchObject([{ ordinal: 1, detail: 'ambiguous: no adjacent text donor' }]);
    const xml = await documentXml(await doc.export());
    const paragraphs = xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? [];
    expect([paragraphs[0], paragraphs[2]]).toEqual([before, after]);
    expect(paragraphs[1]).toContain('First tail');
  } finally { doc.close(); }
}, 30_000);

test('preserves untouched runs across three replacements in one paragraph', async () => {
  const left = run(1, 'Before ');
  const slash = run(3, ' / ', true);
  const plus = run(5, ' + ');
  const right = run(7, ' after.', true);
  await checkProposal([left, run(2, '(name)', true), slash, run(4, '(street)'), plus, run(6, '(city)', true), right], [
    { query: '(name)', newText: 'Ada Lovelace' }, { query: '(street)', newText: 'Elm Street' }, { query: '(city)', newText: 'Leeds city' },
  ], paragraph => expectRuns(paragraph, [left, slash, plus, right], [
    ['Before ', false], ['Ada Lovelace', true], [' / ', true], ['Elm Street', false],
    [' + ', false], ['Leeds city', true], [' after.', true],
  ]));
}, 30_000);

test('reopens the exported paragraph with both replacements', async () => {
  await checkProposal(contractorRuns, contractorEdits, async (paragraph, bytes) => {
    const reopened = await openDocx(bytes);
    try {
      expect(reopened.list().items[1]?.text).toBe('The contractor is Great Lakes Analytical Services, located in Chicago, Illinois.');
    } finally { reopened.close(); }
    expectRuns(paragraph, contractorUntouched, contractorExpected);
  });
}, 30_000);
