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
const run = (text: string) => `<w:r><w:rPr/><w:t>${text}</w:t></w:r>`;
const value = (text: string) => `<w:r><w:rPr/><w:t xml:space="preserve">${text}</w:t></w:r>`;
const middle = run('MRID') + value(' xxxxxxxx');
const replaced = run('MRID') + value(' 12345678');
const empty = '<w:r w:rsidR="12345678"><w:rPr/><w:t/></w:r>';
const tab = '<w:r><w:rPr/><w:t>]</w:t><w:tab/></w:r>';

for (const [name, source, expected] of [
  ['brackets', run('[') + middle + run(']'), replaced],
  ['italic run after prefix', value('Prefix ') + '<w:r><w:rPr><w:i/></w:rPr><w:t>[</w:t></w:r>' + middle + run(']'),
    value('Prefix ') + '<w:r><w:rPr><w:i/></w:rPr><w:t>MRID</w:t></w:r>' +
    '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> 12345678</w:t></w:r>'],
  ['partial final run', run('[') + middle + value('] suffix'), replaced + value(' suffix')],
  ['source empty run', empty + run('[') + middle + run(']'), empty + replaced],
  ['tab beside deleted text', run('[') + middle + tab, replaced + tab.replace('>]<', '><')],
] as const) {
  test(`untracked accept removes emptied text runs and preserves ${name}`, async () => {
    const root = await mkdtemp(join(tmpdir(), 'office-emptied-runs-'));
    const server = await createOfficeMcpServer({ root });
    const client = new Client({ name: 'emptied-runs-test', version: '1' });
    try {
      const kept = '<w:p w:rsidR="87654321"><w:r><w:t>Keep &#x61;ll bytes.</w:t></w:r></w:p>';
      const prefix = `<w:document xmlns:w="${W}"><w:body>${kept}<w:p><w:pPr/>`;
      const suffix = `</w:p>${kept}<w:sectPr/></w:body></w:document>`;
      const zip = new JSZip();
      zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
      zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="main" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
      zip.file('word/document.xml', prefix + source + suffix);
      await writeFile(join(root, 'source.docx'), await zip.generateAsync({ type: 'uint8array' }));
      const [a, b] = InMemoryTransport.createLinkedPair();
      await server.connect(b);
      await client.connect(a);
      const call = async (name: string, args: Record<string, unknown>) => {
        const result = await client.callTool({ name, arguments: args });
        expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true);
        return result.structuredContent as Record<string, any>;
      };
      const { document } = await call('office_open', { path: 'source.docx' });
      const { matches } = await call('office_grep', { document, query: '[MRID xxxxxxxx]' });
      expect(matches).toHaveLength(1);
      const proposal = await call('office_propose', { document, author: 'test', edits: [
        { match: matches[0].match, newText: 'MRID 12345678' },
      ] });
      expect(await call('office_accept', { document, proposal: proposal.id, tracked: false }))
        .toMatchObject({ status: 'accepted', tracked: false });
      await call('office_export', { document, path: 'result.docx' });
      const saved = await JSZip.loadAsync(await readFile(join(root, 'result.docx')));
      expect(await saved.file('word/document.xml')!.async('string')).toBe(prefix + expected + suffix);
    } finally {
      await client.close();
      await server.close();
      await rm(root, { recursive: true, force: true });
    }
  }, 30_000);
}
