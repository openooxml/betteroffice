import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createYrsSession, saveYrsDocx } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

test.each(['split', 'existing'])('%s paragraph hyperlinks resolve relationships under a conflicting source prefix', async (kind) => {
  const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:r="urn:other"><w:body><w:p><w:r><w:t>Host</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)],
  ])));
  const session = await createYrsSession();
  try {
    session.openDocx(bytes, true);
    const host = session.paragraphs('body')[0]!;
    let paraId = host.paraId;
    if (kind === 'split') {
      paraId = session.splitParagraph({ story: 'body', paraId, offset: 4 }).secondParaId;
      session.insertText({ story: 'body', paraId, offset: 0 }, 'Link');
    }
    session.setHyperlink({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 4 } }, { href: 'https://example.com' });
    const saved = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
    const paragraph = xml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*<w:hyperlink\b[^>]*>/)![0];
    const scope = xml.match(/<w:document\b[^>]*>/)![0] + paragraph;
    expect([...scope.matchAll(/xmlns:r="([^"]*)"/g)].at(-1)?.[1]).toBe(R);
    const reopened = await parseDocx(saved.buffer as ArrayBuffer, { preloadFonts: false });
    const links = reopened.package.document.content.flatMap(block => block.type === 'paragraph'
      ? block.content.filter(inline => inline.type === 'hyperlink') : []);
    expect(links).toHaveLength(1);
    expect(links[0]!.href).toBe('https://example.com');
    expect(xml).toContain('>Host</w:t>');
    if (kind === 'split') expect(xml).toContain('>Link</w:t>');
  } finally { session.destroy(); }
});

test('persisting paragraph IDs after a text edit keeps their Word namespace under a conflicting source prefix', async () => {
  const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:w14="urn:other"><w:body><w:p><w:r><w:t>Host</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)],
  ])));
  const session = await createYrsSession();
  try {
    session.openDocx(bytes, true);
    const host = session.paragraphs('body')[0]!;
    session.insertText({ story: 'body', paraId: host.paraId, offset: 0 }, 'Edited ');
    const receipt = session.persistParagraphIds();
    expect(receipt.status).toBe('applied');
    if (receipt.status !== 'applied') throw new Error('paragraph IDs were not persisted');
    const saved = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
    const paragraph = xml.match(/<w:p\b[^>]*>/)![0];
    const scope = xml.match(/<w:document\b[^>]*>/)![0] + paragraph;
    const bindings = new Map([...scope.matchAll(/xmlns:([\w.-]+)="([^"]*)"/g)].map(match => [match[1]!, match[2]!]));
    const attributes = [...paragraph.matchAll(/\b([\w.-]+):paraId="([^"]*)"/g)];
    expect(attributes).toHaveLength(1);
    const ids = attributes.filter(match => bindings.get(match[1]!) === W14);
    expect(ids).toHaveLength(1);
    expect(ids[0]![2]).toBe(receipt.assignments[0]!.ooxmlParaId);
    expect(xml).toContain('xmlns:w14="urn:other"');
    expect(xml).toContain('>Edited Host</w:t>');
    const reopened = await parseDocx(saved.buffer as ArrayBuffer, { preloadFonts: false });
    const written = reopened.package.document.content[0]!;
    if (written.type !== 'paragraph') throw new Error('missing paragraph');
    expect(written.paraId).toBe(ids[0]![2]);
  } finally { session.destroy(); }
});
