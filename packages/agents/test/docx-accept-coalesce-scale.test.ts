import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createOfficeMcpServer } from '../src/mcp';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const B = '<w:b/>';
const I = '<w:i/>';
const C = '<w:color w:val="C00000"/>';
const run = (text: string, rPr: string) => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const SHAPES: { runs: [string, string][]; edits: [string, string][] }[] = [
  { runs: [['Release the belt tension and check ', ''], ['the covers', B], [' before you start.', '']],
    edits: [['Release the belt tension', 'Die Riemenspannung lösen'], ['check the cov', 'die Abdeckungen prü'], ['before you start', 'bevor Sie beginnen']] },
  { runs: [['Tighten ', ''], ['all screws', B], [' on the ', ''], ['left side', I], [' firmly.', '']],
    edits: [['Tighten all screws on the left side firmly', 'Alle Schrauben auf der linken Seite fest anziehen']] },
  { runs: [['Warning: ', C], ['disconnect the power supply', B], [' first.', '']],
    edits: [['Warning: disconnect', 'Warnung: trennen Sie'], ['power supply first', 'Stromversorgung zuerst']] },
  { runs: [['The motor runs at ', ''], ['1500 rpm', B], [' under normal load conditions.', '']],
    edits: [['The motor runs at 15', 'Der Motor läuft mit 15'], ['rpm under normal load conditions', 'U/min unter normaler Last']] },
  { runs: [['Replace the filter every ', ''], ['six months', I], ['.', '']],
    edits: [['Replace the filter every', 'Ersetzen Sie den Filter alle'], ['six months', 'sechs Monate']] },
  { runs: [['Unchanged note ', ''], ['kept', B], [' exactly as it is.', '']], edits: [] },
];

export function scenario() {
  const paragraphs: string[] = [];
  const edits: { paragraph: number; query: string; newText: string }[] = [];
  for (let index = 0; index < 561; index++) {
    const shape = SHAPES[index % SHAPES.length]!;
    const tag = `[${String(index).padStart(3, '0')}] `;
    paragraphs.push(`<w:p>${shape.runs.map(([text, rPr], at) => run(at === 0 ? tag + text : text, rPr)).join('')}</w:p>`);
    shape.edits.forEach(([query, newText], at) => edits.push({ paragraph: index, query: at === 0 ? tag + query : query, newText: at === 0 ? tag + newText : newText }));
  }
  paragraphs.push(`<w:p>${run('Final note ', '')}${run('stays', B)}</w:p>`);
  return { paragraphs, edits, xml: `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraphs.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>` };
}

export async function fixture(xml: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type: 'uint8array' });
}

const paragraphsOf = (xml: string) => xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
const runsOf = (paragraph: string) => [...paragraph.matchAll(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g)].map(match => ({
  rPr: match[1]!.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/)?.[1] ?? '',
  text: [...match[1]!.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(text => text[1]).join(''),
  start: match.index!,
  end: match.index! + match[0].length,
}));
const samePairs = (xml: string) => paragraphsOf(xml).reduce((count, paragraph) => {
  const runs = runsOf(paragraph);
  return count + runs.filter((item, at) => at > 0 && item.rPr === runs[at - 1]!.rPr && item.start === runs[at - 1]!.end && item.text && runs[at - 1]!.text).length;
}, 0);
const chars = (paragraph: string) => runsOf(paragraph).flatMap(item => [...item.text].map(char => [char, item.rPr] as const));

test('span replacements add no same-format fragments to a document of 1,873 runs', async () => {
  const { paragraphs, edits, xml } = scenario();
  const root = await mkdtemp(join(tmpdir(), 'office-coalesce-scale-'));
  const server = await createOfficeMcpServer({ root });
  const client = new Client({ name: 'coalesce-scale-test', version: '1' });
  try {
    await writeFile(join(root, 'source.docx'), await fixture(xml));
    const [a, b] = InMemoryTransport.createLinkedPair();
    await server.connect(b);
    await client.connect(a);
    const call = async (name: string, args: Record<string, unknown>) => {
      const result = await client.callTool({ name, arguments: args });
      if (result.isError) throw new Error(`${name}: ${JSON.stringify(result.structuredContent ?? result.content)}`);
      return result.structuredContent as any;
    };
    const { document } = await call('office_open', { path: 'source.docx' });
    for (let start = 0; start < edits.length; start += 32) {
      const batch = [];
      const used = new Map<string, number>();
      for (const { query, newText } of edits.slice(start, start + 32)) {
        const { matches } = await call('office_grep', { document, query, caseSensitive: true });
        const at = used.get(query) ?? 0;
        used.set(query, at + 1);
        batch.push({ match: matches[at].match, newText });
      }
      const proposal = await call('office_propose', { document, author: 'agent', edits: batch });
      await call('office_review', { document, proposal: proposal.id });
      expect((await call('office_accept', { document, proposal: proposal.id })).status).toBe('accepted');
    }
    await call('office_export', { document, path: 'result.docx' });
    const saved = await (await JSZip.loadAsync(await readFile(join(root, 'result.docx')))).file('word/document.xml')!.async('string');
    const count = (value: string) => value.match(/<w:r\b/g)?.length ?? 0;
    console.log(`RUNS source=${count(xml)} saved=${count(saved)} samePairs source=${samePairs(xml)} saved=${samePairs(saved)}`);
    expect(count(xml)).toBe(1873);
    expect(samePairs(xml)).toBe(0);
    expect(samePairs(saved)).toBe(0);
    expect(count(saved)).toBeLessThanOrEqual(count(xml));
    const out = paragraphsOf(saved);
    expect(out).toHaveLength(paragraphs.length);
    paragraphs.forEach((source, index) => {
      const plan = edits.filter(edit => edit.paragraph === index);
      if (!plan.length) expect(out[index]).toBe(source);
      const from = chars(source);
      const to = chars(out[index]!);
      const fromText = from.map(item => item[0]).join('');
      const toText = to.map(item => item[0]).join('');
      let x = 0;
      let y = 0;
      const same = (length: number) => {
        for (let at = 0; at < length; at++) expect(to[y + at]).toEqual(from[x + at]!);
      };
      for (const { query, newText } of plan) {
        const left = fromText.indexOf(query, x);
        const right = toText.indexOf(newText, y);
        expect(right - y).toBe(left - x);
        same(left - x);
        x = left + query.length;
        y = right + newText.length;
      }
      expect(to.length - y).toBe(from.length - x);
      same(from.length - x);
    });
  } finally {
    await client.close();
    await server.close();
    await rm(root, { recursive: true, force: true });
  }
}, 60_000);
