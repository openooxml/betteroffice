import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { compareDocx } from '../core';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from '../yrs/index';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from './rezip/parts';
import { unzipContainer } from './wasm';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  `xmlns:r="${REL}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;

const paragraph = (start: string, text: string) =>
  `${start}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** A body of paragraphs whose Word IDs repeat, are missing and are lowercase. */
function docx(texts: [string, string, string, string]): Uint8Array {
  const starts = [
    '<w:p w14:paraId="0000ABCD">',
    '<w:p w14:paraId="0000ABCD">',
    '<w:p>',
    '<w:p w14:paraId="0000abcd">',
  ];
  const body = starts.map((start, index) => paragraph(start, texts[index]!)).join('');
  const parts: PartsMap = new Map();
  const set = (name: string, content: string) => parts.set(name, toBytes(content));
  set(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`
  );
  set(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/></Relationships>`
  );
  set(
    'word/styles.xml',
    `<w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`
  );
  set(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function startTags(bytes: Uint8Array): string[] {
  const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  return [...xml.matchAll(/<w:p(?:\s[^>]*)?>/g)].map((match) => match[0]);
}

const sessions: YrsSession[] = [];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

afterEach(() => {
  for (const created of sessions.splice(0)) created.destroy();
});

describe('comparisons and paragraph identities', () => {
  it('keep every source paragraph ID and allocate none', async () => {
    const original = docx(['First', 'Second', 'Third', 'Fourth']);
    const result = await compareDocx(original, docx(['First', 'Second edit', 'Third edit', 'Fourth edit']), {
      author: 'Reviewer',
      date: '2024-05-06T07:08:09Z',
    });
    if (!result.ok) throw new Error(JSON.stringify(result.diagnostics));
    expect(result.changes).toHaveLength(3);
    expect(result.diagnostics.map((diagnostic) => diagnostic.code)).toEqual(['ambiguous-identity']);
    expect(startTags(result.docx)).toEqual(startTags(original));

    const reopened = await createYrsSession({ clientId: 84001 });
    sessions.push(reopened);
    reopened.openDocx(result.docx, true);
    expect(
      reopened
        .paragraphIdentities()
        .paragraphs.filter((identity) => identity.session?.story === 'body')
        .map((identity) => [identity.ooxmlParaId, identity.idOrigin])
    ).toEqual([
      ['0000ABCD', 'source'],
      ['0000ABCD', 'source'],
      [null, null],
      ['0000abcd', 'source'],
    ]);
  });
});
