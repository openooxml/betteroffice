import { expect, spyOn, test } from 'bun:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import JSZip from 'jszip';
import { openDocx } from '../src';
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

async function fixture(runs: readonly string[]): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="doc" Type="${R}/officeDocument" Target="${DOCUMENT}"/></Relationships>`);
  zip.file(DOCUMENT, `<w:document xmlns:w="${W}"><w:body>${before}<w:p>${runs.join('')}</w:p>${after}<w:sectPr/></w:body></w:document>`);
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

test.todo('tracked acceptance wraps only the replaced spans and keeps untouched runs byte-identical', async () => {
  await checkProposal(contractorRuns, contractorEdits, paragraph => {
    const plainRuns: string[] = [];
    const revisionTag = /<w:(ins|del)\b[^>]*>([\s\S]*?)<\/w:\1>/;
    const items = [...paragraph.matchAll(new RegExp(`${revisionTag.source}|${RUN.source}`, 'g'))];
    expect(items.map(([xml, revision, content]) => {
      if (revision) return [revision, runsOf(content!).map(xml => runText(xml, revision === 'del' ? 'delText' : 't'))];
      plainRuns.push(xml);
      return runText(xml);
    }), paragraph).toEqual([
      ['The contractor is ', false], ['del', [[contractor.query, true]]], ['ins', [[contractor.newText, true]]],
      [', ', true], ['located', false], [' ', false], ['in ', false],
      ['del', [[city.query, true]]], ['ins', [[city.newText, true]]], ['.', true],
    ]);
    expect(paragraph.match(/<w:(ins|del)\b/g)).toHaveLength(4);
    for (const xml of contractorUntouched) expect(plainRuns).toContain(xml);
  }, { tracked: true, date: '2026-01-01T00:00:00Z' });
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
