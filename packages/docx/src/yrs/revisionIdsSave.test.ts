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

function docx(body: string): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  set(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set(DOCUMENT, `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (...content: string[]) => `<w:p>${content.join('')}</w:p>`;

/** Every `w:id` an annotation start element of `xml` holds, by element name. */
function annotationIds(xml: string): Array<[string, number]> {
  return [...xml.matchAll(/<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|bookmarkStart|commentRangeStart)\b[^>]*\bw:id="(-?\d+)"/g)]
    .map(([, name, id]) => [name!, Number(id)]);
}

const sessions: YrsSession[] = [];

async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: 3 });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

function replace(session: YrsSession, text: string, from: string, to: string) {
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
  it('gives a tracked replace distinct deletion and insertion ids in Word range', async () => {
    const session = await open(docx(p(run('The cat sat.')) + p(run('The dog ran.'))));
    replace(session, 'The cat sat.', 'cat', 'lion');
    replace(session, 'The dog ran.', 'dog', 'wolf');
    const ids = annotationIds((await save(session)).xml);

    expect(ids.map(([name]) => name).sort()).toEqual(['del', 'del', 'ins', 'ins']);
    expect(new Set(ids.map(([, id]) => id)).size).toBe(4);
    for (const [, id] of ids) {
      expect(id).toBeGreaterThan(0);
      expect(id).toBeLessThan(2 ** 31);
    }
  });

  it('keeps source annotation ids and numbers new revisions above them', async () => {
    const source = docx(
      p(
        `<w:ins w:id="7" w:author="A" w:date="2025-01-01T00:00:00Z">${run('Added ')}</w:ins>`,
        run('The cat sat.')
      ) +
        p('<w:bookmarkStart w:id="12" w:name="mark"/>', run('Marked.'), '<w:bookmarkEnd w:id="12"/>') +
        p(`<w:del w:id="3" w:author="A" w:date="2025-01-01T00:00:00Z"><w:r><w:delText>Old </w:delText></w:r></w:del>`, run('Kept.'))
    );
    const session = await open(source);
    replace(session, 'The cat sat.', 'cat', 'lion');
    const ids = annotationIds((await save(session)).xml);

    expect(ids).toContainEqual(['ins', 7]);
    expect(ids).toContainEqual(['bookmarkStart', 12]);
    expect(ids).toContainEqual(['del', 3]);
    const added = ids.filter(([, id]) => ![3, 7, 12].includes(id));
    expect(added.map(([name]) => name).sort()).toEqual(['del', 'ins']);
    for (const [, id] of added) expect(id).toBeGreaterThan(12);
    expect(new Set(ids.map(([, id]) => id)).size).toBe(ids.length);
  });

  it('renumbers a source revision id that an adjacent replace reuses for its deletion', async () => {
    const session = await open(
      docx(
        p(
          `<w:ins w:id="7" w:author="Reviewer" w:date="2025-01-01T00:00:00Z">${run('Added ')}</w:ins>`,
          run('The cat sat.')
        )
      )
    );
    replace(session, 'The cat sat.', 'The', 'A');
    const ids = annotationIds((await save(session)).xml);
    const deletions = ids.filter(([name]) => name === 'del');

    expect(ids[0]).toEqual(['ins', 7]);
    expect(deletions).toHaveLength(1);
    expect(ids.filter(([, id]) => id === deletions[0]![1])).toEqual(deletions);
    for (const [, id] of ids) expect(id).toBeLessThan(2 ** 31);
  });

  it('pairs a saved replace for accept and reject after reopening', async () => {
    const session = await open(docx(p(run('The cat sat.'))));
    const receipt = replace(session, 'The cat sat.', 'cat', 'lion');
    const { bytes } = await save(session);

    const live = session.listRevisions();
    expect(live.map(({ kind }) => kind).sort()).toEqual(['deletion', 'insertion']);
    expect(new Set(live.map(({ revisionId }) => revisionId))).toEqual(new Set([receipt.revisionId!]));

    for (const resolution of ['accept', 'reject'] as const) {
      const reopened = await open(bytes);
      const revisions = reopened.listRevisions();
      const deletion = revisions.find(({ kind }) => kind === 'deletion')!;
      const insertion = revisions.find(({ kind }) => kind === 'insertion')!;
      expect(revisions).toHaveLength(2);
      expect(deletion.revisionId).not.toBe(insertion.revisionId);
      expect([deletion.author, deletion.date]).toEqual([insertion.author, insertion.date]);
      expect(deletion.range.end).toEqual(insertion.range.start);
      for (const { revisionId } of revisions) {
        if (resolution === 'accept') reopened.acceptChange({ revisionId });
        else reopened.rejectChange({ revisionId });
      }
      expect(reopened.listRevisions()).toEqual([]);
      expect(texts(reopened)).toEqual([resolution === 'accept' ? 'The lion sat.' : 'The cat sat.']);
    }

    session.acceptChange({ revisionId: receipt.revisionId! });
    expect(session.listRevisions()).toEqual([]);
    expect(texts(session)).toEqual(['The lion sat.']);
  });
});
