import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';
import { captureSessionSave, writeSessionSave } from './saveYrsDocx';
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

function paragraph(session: YrsSession, text: string) {
  const found = session.paragraphs('body').find((candidate) => candidate.text === text);
  if (!found) throw new Error(`no paragraph reads ${text}`);
  return found;
}

function structuralFixture() {
  const texts = [
    'Before QA',
    'Split QA',
    'Kept QA',
    'Merge QA',
    'Into QA',
    'Middle QA',
    'Insert QA',
    'Remove QA',
    'After QA',
  ];
  const paragraphs = texts.map((text, index) => {
    const id = (index + 1).toString(16).padStart(8, '0').toUpperCase();
    return `<w:p w14:paraId="${id}" w:rsidR="00AA0001"><w:pPr><w:jc w:val="center"/><w:spacing w:after="${index * 20}"/></w:pPr><w:r w:rsidRPr="00AA0002"><w:t>${text}</w:t></w:r></w:p>`;
  });
  return {
    texts,
    paragraphs,
    bytes: docx(paragraphs.join('\n'), `xmlns:w="${W}" xmlns:w14="${W14}"`),
  };
}

async function equivalentSave(session: YrsSession, untouched: string[], texts: string[]) {
  const base = session.materializeDocx()!;
  const capture = captureSessionSave(session);
  const document = yrsToDocument(session, base);
  const options = { updateModifiedDate: false };
  const saved = await saveYrsDocx(session, options);
  const whole = await writeSessionSave(
    session,
    document,
    capture,
    base.originalBuffer!,
    options,
    () => false
  );
  const xml = part(saved.bytes);
  const wholeXml = part(whole.bytes);
  for (const source of untouched) {
    expect(xml).toContain(source);
    expect(wholeXml).not.toContain(source);
  }
  const reopened = await open(saved.bytes);
  const full = await open(whole.bytes);
  const paragraphs = (session: YrsSession) =>
    session.paragraphs('body').map(({ text, properties }) => ({ text, properties }));
  expect(reopened.paragraphs('body').map(({ text }) => text)).toEqual(texts);
  expect(paragraphs(reopened)).toEqual(paragraphs(full));
  return xml;
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

  it('writes split, merged and appended paragraphs and keeps every other paragraph as authored', async () => {
    const kept = '<w:p w14:paraId="00000002" w:rsidR="00AA0001"><w:r w:rsidRPr="00AA0002"><w:t>Kept QA</w:t></w:r></w:p>';
    const last = '<w:p w14:paraId="00000005" w:rsidR="00AA0003"><w:r><w:t>Last QA</w:t></w:r></w:p>';
    const section = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';
    const source = docx(
      [
        '<w:p w14:paraId="00000001"><w:r><w:t>First QA</w:t></w:r></w:p>',
        kept,
        '<w:p w14:paraId="00000003"><w:r><w:t>Merge QA</w:t></w:r></w:p>',
        '<w:p w14:paraId="00000004"><w:r><w:t>Into QA</w:t></w:r></w:p>',
        last,
        section,
      ].join(''),
      `xmlns:w="${W}" xmlns:w14="${W14}"`
    );
    const session = await open(source);
    const paragraph = (text: string) => {
      const found = session.paragraphs('body').find((candidate) => candidate.text === text);
      if (!found) throw new Error(`no paragraph reads ${text}`);
      return found;
    };
    session.splitParagraph({ story: 'body', paraId: paragraph('First QA').paraId, offset: 5 });
    session.mergeParagraphs('body', paragraph('Merge QA').paraId);
    session.splitParagraph({ story: 'body', paraId: paragraph('Last QA').paraId, offset: 7 });
    const bytes = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    const xml = part(bytes);

    expect(xml).toContain(`${kept}<w:p`);
    expect(xml).toContain(last);
    expect(xml.endsWith(`${section}</w:body></w:document>`)).toBe(true);
    expect(paragraphsWith(xml, 'Merge QA')).toHaveLength(1);
    expect(paragraphsWith(xml, 'Into QA')).toHaveLength(1);
    const reopened = await open(bytes);
    expect(reopened.paragraphs('body').map(({ text }) => text)).toEqual([
      'First',
      ' QA',
      'Kept QA',
      'Merge QAInto QA',
      'Last QA',
      '',
    ]);
  });

  it('keeps the paragraphs around a tracked paragraph insertion as authored', async () => {
    const before = '<w:p w14:paraId="00000001" w:rsidR="00AA0001"><w:r><w:t>Before QA</w:t></w:r></w:p>';
    const after = '<w:p w14:paraId="00000002" w:rsidR="00AA0002"><w:r><w:t>After QA</w:t></w:r></w:p>';
    const source = docx(`${before}${after}`, `xmlns:w="${W}" xmlns:w14="${W14}"`);
    const session = await open(source);
    const [first] = session.paragraphs('body');
    const author = { name: 'Reviewer', date: '2026-01-01T00:00:00Z' };
    session.splitParagraph({ story: 'body', paraId: first!.paraId, offset: 'Before QA'.length }, author);
    const inserted = session.paragraphs('body')[1]!;
    session.insertText({ story: 'body', paraId: inserted.paraId, offset: 0 }, 'New QA', author);
    const bytes = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    const xml = part(bytes);

    expect(xml).toContain(after);
    expect(paragraphsWith(xml, 'New QA')).toHaveLength(1);
    expect(paragraphsWith(xml, 'New QA')[0]).toContain('<w:ins ');
    const reopened = await open(bytes);
    expect(reopened.paragraphs('body').map(({ text }) => text)).toEqual([
      'Before QA',
      'New QA',
      'After QA',
    ]);
  });

  for (const edit of ['split', 'merge', 'insert', 'remove', 'combined'] as const) {
    it(`matches the whole-part save after ${edit} and keeps every untouched paragraph verbatim`, async () => {
      const fixture = structuralFixture();
      const session = await open(fixture.bytes);
      const untouched = [...fixture.paragraphs];
      const expected = [...fixture.texts];
      const forget = (...texts: string[]) => {
        for (const text of texts) {
          const source = fixture.paragraphs[fixture.texts.indexOf(text)]!;
          untouched.splice(untouched.indexOf(source), 1);
        }
      };
      if (edit === 'split' || edit === 'combined') {
        session.splitParagraph({
          story: 'body',
          paraId: paragraph(session, 'Split QA').paraId,
          offset: 5,
        });
        expected.splice(expected.indexOf('Split QA'), 1, 'Split', ' QA');
        forget('Split QA');
      }
      if (edit === 'merge' || edit === 'combined') {
        session.mergeParagraphs('body', paragraph(session, 'Merge QA').paraId);
        expected.splice(expected.indexOf('Merge QA'), 2, 'Merge QAInto QA');
        forget('Merge QA', 'Into QA');
      }
      if (edit === 'insert' || edit === 'combined') {
        const split = session.splitParagraph({
          story: 'body',
          paraId: paragraph(session, 'Insert QA').paraId,
          offset: 'Insert QA'.length,
        });
        session.insertText({ story: 'body', paraId: split.secondParaId, offset: 0 }, 'New QA');
        expected.splice(expected.indexOf('Insert QA') + 1, 0, 'New QA');
      }
      if (edit === 'remove' || edit === 'combined') {
        const paragraphs = session.paragraphs('body');
        const at = paragraphs.findIndex(({ text }) => text === 'Remove QA');
        const before = paragraphs[at - 1]!;
        const removed = paragraphs[at]!;
        session.deleteRange({
          story: 'body',
          start: { paraId: before.paraId, offset: before.text.length },
          end: { paraId: removed.paraId, offset: removed.text.length },
        });
        expected.splice(expected.indexOf('Remove QA'), 1);
        forget('Remove QA');
      }
      await equivalentSave(session, untouched, expected);
    });
  }

  for (const edit of ['insertion', 'removal'] as const) {
    for (const resolution of ['pending', 'accept', 'reject'] as const) {
      it(`keeps untouched paragraphs through a tracked paragraph ${edit} and ${resolution}`, async () => {
        const fixture = structuralFixture();
        const session = await open(fixture.bytes);
        const author = { name: 'Reviewer', date: '2026-01-01T00:00:00Z' };
        const expected = [...fixture.texts];
        const revisions: string[] = [];
        let marked: string;
        let untouched: string[];
        if (edit === 'insertion') {
          const split = session.splitParagraph(
            {
              story: 'body',
              paraId: paragraph(session, 'Insert QA').paraId,
              offset: 'Insert QA'.length,
            },
            author
          );
          const text = session.insertText(
            { story: 'body', paraId: split.secondParaId, offset: 0 },
            'New QA',
            author
          );
          expect(split.revisionId).not.toBeNull();
          expect(text.revisionId).not.toBeNull();
          revisions.push(text.revisionId!, split.revisionId!);
          expected.splice(expected.indexOf('Insert QA') + 1, 0, 'New QA');
          marked = 'Insert QA';
          untouched = fixture.paragraphs.filter((_, index) => fixture.texts[index] !== marked);
        } else {
          const receipt = session.deleteRange(
            {
              story: 'body',
              start: { paraId: paragraph(session, 'Remove QA').paraId, offset: 0 },
              end: { paraId: paragraph(session, 'After QA').paraId, offset: 0 },
            },
            author
          );
          expect(receipt.revisionId).not.toBeNull();
          revisions.push(receipt.revisionId!);
          marked = 'Remove QA';
          untouched = fixture.paragraphs.filter((_, index) => fixture.texts[index] !== marked);
        }
        const xml = await equivalentSave(session, untouched, expected);
        expect(paragraphsWith(xml, marked)[0]).toMatch(
          edit === 'insertion'
            ? /<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:ins /
            : /<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:del /
        );
        if (resolution === 'pending') return;
        for (const revisionId of new Set(revisions)) {
          if (resolution === 'accept') session.acceptChange({ revisionId });
          else session.rejectChange({ revisionId });
        }
        if (edit === 'insertion' && resolution === 'reject') {
          expected.splice(expected.indexOf('New QA'), 1);
        }
        if (edit === 'removal' && resolution === 'accept') {
          expected.splice(expected.indexOf('Remove QA'), 1);
        }
        await equivalentSave(session, untouched, expected);
      });
    }
  }

  it(
    'keeps inherited paragraph properties inherited in the edited paragraph (#1067)',
    async () => {
      const session = await open(inheritedSpacing());
      insertAtStart(session, 'Conclusion QA', 'Edited ');
      const xml = part((await saveYrsDocx(session)).bytes);
      expect(paragraphsWith(xml, 'Edited Conclusion QA')).toEqual([
        '<w:p><w:r><w:t>Edited Conclusion QA</w:t></w:r></w:p>',
      ]);
    }
  );
  it(
    'keeps inherited paragraph properties inherited through yrsToDocument and repackDocx (#1067)',
    async () => {
      const session = await open(inheritedSpacing());
      const xml = part(
        new Uint8Array(await repackDocx(yrsToDocument(session, session.materializeDocx()!)))
      );
      expect(paragraphsWith(xml, 'Finding QA')).toEqual(['<w:p><w:r><w:t>Finding QA</w:t></w:r></w:p>']);
    }
  );
});
