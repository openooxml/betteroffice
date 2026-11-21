import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const AFTER = 'After the contents';
const TYPED = 'Typed';
const sessions: YrsSession[] = [];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const fieldCharacter = (kind: string) => `<w:r><w:fldChar w:fldCharType="${kind}"/></w:r>`;
const paragraph = (id: string, content: string) => `<w:p w14:paraId="${id}">${content}</w:p>`;

/** A table of contents whose field begins in its first entry's paragraph and ends in its last, then one paragraph after it. */
const PARAGRAPHS = [
  paragraph('0000001A', fieldCharacter('begin') +
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
    fieldCharacter('separate') + run('Introduction')),
  paragraph('0000001B', run('Scope')),
  paragraph('0000001C', run('Terms')),
  paragraph('0000001D', run('Appendix') + fieldCharacter('end')),
  paragraph('0000001E', run(AFTER)),
];
const BODY = PARAGRAPHS.join('');

function fixture(): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${BODY}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** The begin, separate and end field characters of `xml`, the lowest running depth of open fields and where the last end is. */
function fieldCharacters(xml: string): { counts: [number, number, number]; lowest: number; lastEnd: number } {
  const counts: [number, number, number] = [0, 0, 0];
  let depth = 0;
  let lowest = 0;
  let lastEnd = -1;
  for (const match of xml.matchAll(/w:fldCharType="([^"]*)"/g)) {
    if (match[1] === 'begin') {
      counts[0] += 1;
      depth += 1;
    } else if (match[1] === 'separate') {
      counts[1] += 1;
    } else if (match[1] === 'end') {
      counts[2] += 1;
      depth -= 1;
      lastEnd = match.index ?? -1;
    } else {
      throw new Error(`field character of type ${match[1]}`);
    }
    lowest = Math.min(lowest, depth);
  }
  return { counts, lowest, lastEnd };
}

for (const [name, index, offset] of [['closing', 3, 1], ['opening', 0, 0], ['middle', 1, 1]] as const) {
  test(`typing into the ${name} paragraph of a field spanning paragraphs saves the field balanced`, async () => {
    const session = await createYrsSession({ clientId: 98100 + sessions.length });
    sessions.push(session);
    session.seedFromDocx(fixture());
    const paragraphs = session.paragraphs('body');
    expect(paragraphs.length).toBe(5);
    session.insertText({ story: 'body', paraId: paragraphs[index]!.paraId, offset }, TYPED);
    const saved = await saveYrsDocx(session, { updateModifiedDate: false });
    const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
    const { counts, lowest, lastEnd } = fieldCharacters(xml);
    expect(counts).toEqual([1, 1, 1]);
    expect(lowest).toBe(0);
    expect(lastEnd).toBeLessThan(xml.indexOf(AFTER));
    expect(xml).toContain(TYPED);
  });
}

/** The entries that do not lie in `xml` after its first field separate and before the first field end after that separate. */
function entriesOutsideTheField(xml: string, entries: readonly string[]): string[] {
  const separate = xml.indexOf('w:fldCharType="separate"');
  const end = separate < 0 ? -1 : xml.indexOf('w:fldCharType="end"', separate);
  return entries.filter((entry) => {
    const at = xml.indexOf(entry);
    return separate < 0 || end < 0 || at < separate || at > end;
  });
}

for (const [name, index, offset, entries] of [
  ['closing', 3, 1, ['Introduction', 'Scope', 'Terms', TYPED]],
  ['opening', 0, 0, ['Introduction', 'Scope', 'Terms', 'Appendix']],
] as const) {
  test(`typing into the ${name} paragraph of a field spanning paragraphs keeps the field spanning them`, async () => {
    const session = await createYrsSession({ clientId: 98200 + sessions.length });
    sessions.push(session);
    session.seedFromDocx(fixture());
    const paragraphs = session.paragraphs('body');
    expect(paragraphs.length).toBe(5);
    session.insertText({ story: 'body', paraId: paragraphs[index]!.paraId, offset }, TYPED);
    const saved = await saveYrsDocx(session, { updateModifiedDate: false });
    const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
    expect(xml).toContain(TYPED);
    for (const [at, kept] of PARAGRAPHS.entries()) {
      if (at !== index) expect(xml).toContain(kept);
    }
    expect(entriesOutsideTheField(xml, entries)).toEqual([]);
  });
}
