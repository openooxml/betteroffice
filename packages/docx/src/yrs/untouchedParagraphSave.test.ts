import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const DOCUMENT = 'word/document.xml';
const decoder = new TextDecoder();

function docx(body: string, namespaces = `xmlns:w="${W}"`): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
  );
  set(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set(
    'word/_rels/document.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`
  );
  set(
    'word/styles.xml',
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`
  );
  set(DOCUMENT, `<w:document ${namespaces}><w:body>${body}</w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function part(bytes: Uint8Array, name = DOCUMENT): string {
  return decoder.decode(unzipContainer(bytes)[name]);
}

/** The `w:p` elements of `xml` that hold `text`, as written. */
function paragraphsWith(xml: string, text: string): string[] {
  return (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).filter((paragraph) =>
    paragraph.includes(text)
  );
}

/** Paragraphs inheriting docDefaults spacing around one that overrides it. */
function inheritedSpacing(): Uint8Array {
  return docx(
    [
      '<w:p><w:r><w:t>Finding QA</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Explicit override QA</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>Conclusion QA</w:t></w:r></w:p>',
    ].join('')
  );
}

const sessions: YrsSession[] = [];

async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: 3 });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

function insertAtStart(session: YrsSession, text: string, inserted: string): void {
  const target = session.paragraphs('body').find((paragraph) => paragraph.text.includes(text));
  if (!target) throw new Error(`no paragraph holds ${text}`);
  session.insertText({ story: 'body', paraId: target.paraId, offset: 0 }, inserted);
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  )
);

afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

describe('a session save after an edit', () => {
  it('re-serializes only the edited paragraph and keeps every other byte of the part', async () => {
    const edited = '<w:p w14:paraId="00000004"><w:r><w:t>Conclusion QA</w:t></w:r></w:p>';
    const body = [
      '<w:p w14:paraId="00000001" w:rsidR="00AA0001" w:rsidRDefault="00AA0001"><w:r w:rsidRPr="00AA0002"><w:t>Finding QA</w:t></w:r></w:p>',
      '<w:p w14:paraId="00000002"><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Explicit override</w:t></w:r></w:p>',
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr><w:p w14:paraId="00000003"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      '<w:p w14:paraId="00000005"><w:bookmarkStart w:id="0" w:name="mark"/><w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="0"/><w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve"> inserted</w:t></w:r></w:ins></w:p>',
      edited,
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
    ].join('');
    const source = docx(body, `xmlns:w="${W}" xmlns:w14="${W14}"`);
    const xml = part(source);
    const start = xml.indexOf(edited);
    const end = start + edited.length;

    const session = await open(source);
    insertAtStart(session, 'Conclusion QA', 'Edited ');
    const saved = part((await saveYrsDocx(session, { updateModifiedDate: false })).bytes);

    expect(saved.slice(0, start)).toBe(xml.slice(0, start));
    expect(saved.endsWith(xml.slice(end))).toBe(true);
    expect(paragraphsWith(saved, 'Edited Conclusion QA')).toHaveLength(1);
  });

  it('keeps inherited paragraph properties inherited and explicit ones as authored across three cycles (#1067)', async () => {
    const source = inheritedSpacing();
    let bytes = source;
    for (let cycle = 1; cycle <= 3; cycle += 1) {
      const session = await open(bytes);
      insertAtStart(session, 'Conclusion QA', `${cycle}`);
      bytes = (await saveYrsDocx(session)).bytes;
      const xml = part(bytes);
      expect(paragraphsWith(xml, 'Finding QA')).toEqual(['<w:p><w:r><w:t>Finding QA</w:t></w:r></w:p>']);
      expect(paragraphsWith(xml, 'Explicit override QA')).toEqual([
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Explicit override QA</w:t></w:r></w:p>',
      ]);
      expect(paragraphsWith(xml, 'Conclusion QA')).toHaveLength(1);
    }
    expect(part(bytes, 'word/styles.xml')).toBe(part(source, 'word/styles.xml'));
  });

  it('still saves a part whose paragraphs were split', async () => {
    const source = docx(
      '<w:p><w:r><w:t>First QA</w:t></w:r></w:p><w:p><w:r><w:t>Second QA</w:t></w:r></w:p>'
    );
    const session = await open(source);
    const [first] = session.paragraphs('body');
    session.splitParagraph({ story: 'body', paraId: first!.paraId, offset: 5 });
    const xml = part((await saveYrsDocx(session)).bytes);
    expect(paragraphsWith(xml, 'First')).toHaveLength(1);
    expect(paragraphsWith(xml, 'QA')).toHaveLength(2);
  });

  it.todo(
    'keeps inherited paragraph properties inherited in the edited paragraph (#1067, deferred-after-0.4.1: touched paragraphs)',
    async () => {
      const session = await open(inheritedSpacing());
      insertAtStart(session, 'Conclusion QA', 'Edited ');
      const xml = part((await saveYrsDocx(session)).bytes);
      expect(paragraphsWith(xml, 'Edited Conclusion QA')).toEqual([
        '<w:p><w:r><w:t>Edited Conclusion QA</w:t></w:r></w:p>',
      ]);
    }
  );
  it.todo(
    'keeps inherited paragraph properties inherited through yrsToDocument and repackDocx (#1067, deferred-after-0.4.1: full repack)',
    async () => {
      const session = await open(inheritedSpacing());
      const xml = part(
        new Uint8Array(await repackDocx(yrsToDocument(session, session.materializeDocx()!)))
      );
      expect(paragraphsWith(xml, 'Finding QA')).toEqual(['<w:p><w:r><w:t>Finding QA</w:t></w:r></w:p>']);
    }
  );
});
