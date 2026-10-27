import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DOCUMENT = 'word/document.xml';
const AUTHOR = { name: 'Reviewer', date: '2026-01-01T00:00:00Z' };
const decoder = new TextDecoder();

function docx(body: string, comments?: string): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set(
    '[Content_Types].xml',
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`
  );
  set(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set(DOCUMENT, `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  if (comments) {
    set(
      'word/_rels/document.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="c" Type="${R}/comments" Target="comments.xml"/></Relationships>`
    );
    set('word/comments.xml', `<w:comments xmlns:w="${W}">${comments}</w:comments>`);
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (...content: string[]) => `<w:p>${content.join('')}</w:p>`;
const SOURCE = 'w:author="Reviewer" w:date="2025-01-01T00:00:00Z"';
const comment = (id: number, text: string) =>
  `<w:commentRangeStart w:id="${id}"/>${run(text)}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
const commentBody = (id: number) =>
  `<w:comment w:id="${id}" w:author="A" w:date="2025-01-01T00:00:00Z"><w:p>${run('Note')}</w:p></w:comment>`;

/** Every `w:id` an annotation start element of `xml` holds, by element name. */
function annotationIds(xml: string): Array<[string, number]> {
  return [...xml.matchAll(/<w:(ins|del|moveFrom|moveTo|pPrChange|bookmarkStart|commentRangeStart)\b[^>]*\bw:id="(-?\d+)"/g)]
    .map(([, name, id]) => [name!, Number(id)]);
}

const sessions: YrsSession[] = [];

async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: 3 });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

function replace(session: YrsSession, from: string, to: string) {
  const target = session.paragraphs('body').find((paragraph) => paragraph.text.includes(from));
  if (!target) throw new Error(`no paragraph holds ${from}`);
  const offset = target.text.indexOf(from);
  return session.replaceRange(
    {
      story: 'body',
      start: { paraId: target.paraId, offset },
      end: { paraId: target.paraId, offset: offset + from.length },
    },
    to,
    AUTHOR
  );
}

async function save(session: YrsSession): Promise<{ bytes: Uint8Array; xml: string }> {
  const { bytes } = await saveYrsDocx(session, { updateModifiedDate: false });
  return { bytes, xml: decoder.decode(unzipContainer(bytes)[DOCUMENT]) };
}

const texts = (session: YrsSession) => session.paragraphs('body').map(({ text }) => text);

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  )
);

afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

describe('saved tracked change ids', () => {
  it('gives each tracked replace one id in Word range, shared by its deletion and insertion', async () => {
    const session = await open(docx(p(run('The cat sat.')) + p(run('The dog ran.'))));
    replace(session, 'cat', 'lion');
    replace(session, 'dog', 'wolf');
    const ids = annotationIds((await save(session)).xml);

    expect(ids.map(([name]) => name)).toEqual(['del', 'ins', 'del', 'ins']);
    expect(ids[0]![1]).toBe(ids[1]![1]);
    expect(ids[2]![1]).toBe(ids[3]![1]);
    expect(ids[0]![1]).not.toBe(ids[2]![1]);
    for (const [, id] of ids) {
      expect(id).toBeGreaterThan(0);
      expect(id).toBeLessThan(2 ** 31);
    }
  });

  it('keeps source ids and numbers new revisions past every annotation id', async () => {
    const source = docx(
      p(`<w:ins w:id="7" ${SOURCE}>${run('Added ')}</w:ins>`, run('The cat sat.')) +
        p('<w:bookmarkStart w:id="12" w:name="mark"/>', run('Marked.'), '<w:bookmarkEnd w:id="12"/>') +
        p(`<w:del w:id="3" ${SOURCE}><w:r><w:delText>Old </w:delText></w:r></w:del>`, run('Kept.')) +
        p(comment(13, 'Noted.'), comment(50, 'Later.')),
      commentBody(13) + commentBody(50)
    );
    const session = await open(source);
    replace(session, 'cat', 'lion');
    const ids = annotationIds((await save(session)).xml);

    for (const kept of [['ins', 7], ['bookmarkStart', 12], ['del', 3], ['commentRangeStart', 13], ['commentRangeStart', 50]]) {
      expect(ids).toContainEqual(kept as [string, number]);
    }
    const added = ids.filter(([, id]) => ![3, 7, 12, 13, 50].includes(id));
    expect(added).toEqual([['del', 14], ['ins', 14]]);
  });

  it('uses the last id in Word range', async () => {
    const session = await open(docx(p(`<w:ins w:id="2147483646" ${SOURCE}>${run('Added ')}</w:ins>`, run('The cat sat.'))));
    replace(session, 'cat', 'lion');

    expect(annotationIds((await save(session)).xml)).toEqual([
      ['ins', 2147483646],
      ['del', 2147483647],
      ['ins', 2147483647],
    ]);
  });

  it('wraps to the lowest free id past the last one in Word range', async () => {
    const session = await open(docx(p(`<w:ins w:id="2147483647" ${SOURCE}>${run('Added ')}</w:ins>`, run('The cat sat.'))));
    replace(session, 'cat', 'lion');

    expect(annotationIds((await save(session)).xml)).toEqual([
      ['ins', 2147483647],
      ['del', 1],
      ['ins', 1],
    ]);
  });

  for (const resolution of ['accept', 'reject'] as const) {
    it(`groups a saved replace after reopening and ${resolution}s it in one undo step`, async () => {
      const session = await open(docx(p(run('The cat sat.'))));
      replace(session, 'cat', 'lion');
      const reopened = await open((await save(session)).bytes);
      const revisions = reopened.listRevisions();

      expect(revisions.map(({ kind }) => kind)).toEqual(['deletion', 'insertion']);
      expect(revisions[0]!.revisionId).toBe(revisions[1]!.revisionId);
      reopened.beginUndoCapture();
      const target = { revisionId: revisions[0]!.revisionId };
      if (resolution === 'accept') reopened.acceptChange(target);
      else reopened.rejectChange(target);
      expect(reopened.listRevisions()).toEqual([]);
      expect(texts(reopened)).toEqual([resolution === 'accept' ? 'The lion sat.' : 'The cat sat.']);
      expect(reopened.undo()).toBe(true);
      expect(reopened.listRevisions()).toEqual(revisions);
      expect(reopened.undo()).toBe(false);
    });
  }

  it('keeps a source paragraph change and the typing that joins it as one revision', async () => {
    const session = await open(
      docx(`<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="7" ${SOURCE}><w:pPr/></w:pPrChange></w:pPr></w:p>`)
    );
    const [paragraph] = session.paragraphs('body');
    session.insertText({ story: 'body', paraId: paragraph!.paraId, offset: 0 }, 'Typed', AUTHOR);
    const { bytes, xml } = await save(session);
    const reopened = await open(bytes);

    expect(annotationIds(xml)).toEqual([['pPrChange', 7], ['ins', 7]]);
    expect(new Set(reopened.listRevisions().map(({ revisionId }) => revisionId)).size).toBe(1);
  });

  it('saves a suggested paragraph property change with an id in Word range', async () => {
    const session = await open(docx(p(run('Body.'))));
    const [paragraph] = session.paragraphs('body');
    const at = { paraId: paragraph!.paraId, offset: 0 };
    session.setParagraphAttrs({ story: 'body', start: at, end: at }, { alignment: 'right' }, AUTHOR);
    session.insertText({ story: 'body', ...at }, 'New ', AUTHOR);
    const live = new Set(session.listRevisions().map(({ revisionId }) => revisionId));
    const { bytes, xml } = await save(session);
    const ids = annotationIds(xml);

    expect(ids.map(([name]) => name).sort()).toEqual(['ins', 'pPrChange']);
    for (const [, id] of ids) expect(id).toBeLessThan(2 ** 31);
    expect(new Set(ids.map(([, id]) => id)).size).toBe(live.size);
    const reopened = await open(bytes);
    const change = reopened.listRevisions().find(({ kind }) => kind === 'pPrChange');
    expect(change?.author).toBe(AUTHOR.name);
    reopened.rejectChange({ revisionId: change!.revisionId });
    expect((await save(reopened)).xml).not.toContain('w:val="right"');
  });
});
