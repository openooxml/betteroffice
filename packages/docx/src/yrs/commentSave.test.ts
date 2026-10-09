import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import {
  captureSessionSave,
  createYrsSession,
  saveYrsDocx,
  writeSessionSave,
  yrsToDocument,
  type YrsSession,
} from './index';
import type { Comment, Paragraph, Run } from '../types/document';
import { documentToYrs } from './documentToYrs';

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);

const sessions: YrsSession[] = [];
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const BODY =
  `<w:p w14:paraId="0C000001"><w:commentRangeStart w:id="1"/>${t('Intro')}<w:commentRangeEnd w:id="1"/>` +
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>${t(' paragraph')}</w:p>` +
  `<w:p w14:paraId="0C000002">${t('See ')}<w:hyperlink r:id="rIdH">${t('the linked text')}</w:hyperlink>` +
  `${t(' then ')}<w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">${t('inserted')}</w:ins>${t(', ')}` +
  `<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${t('control')}</w:sdtContent></w:sdt>` +
  `<w:fldSimple w:instr=" PAGE ">${t('1')}</w:fldSimple>${t(' tail')}</w:p>` +
  `<w:p w14:paraId="0C000003">${t('Closing words')}</w:p>` +
  `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="0C000004">${t('Cell')}</w:p></w:tc></w:tr></w:tbl>` +
  `<w:p w14:paraId="0C000005">${t('After the table')}</w:p>` +
  `<w:p w14:paraId="0C000006">${t('Go to ')}<w:hyperlink w:anchor="target">${t('page ')}` +
  `<w:fldSimple w:instr=" PAGE ">${t('7')}</w:fldSimple>${t(' of the text')}</w:hyperlink>${t(' now')}</w:p>`;

function docx(body = BODY, commentIds = [1]): Uint8Array {
  const parts: Record<string, string> = {
    '[Content_Types].xml':
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdH" Type="${REL}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`,
    'word/comments.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${NS}>${commentIds.map((id) => `<w:comment w:id="${id}" w:author="Reviewer"><w:p w14:paraId="0D00000${id}">${t('Check')}</w:p></w:comment>`).join('')}</w:comments>`,
  };
  return new Uint8Array(
    rezipPartsToArrayBuffer(new Map(Object.entries(parts).map(([name, xml]) => [name, toBytes(xml)])))
  );
}

const IDENTITIES = resolve(
  import.meta.dir,
  '../../../../crates/docx-edit/tests/fixtures/paragraph-identities'
);

/** The shared paragraph identity fixture: a commented body, a header and a footnote. */
function identities(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (dir: string, prefix: string) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      if (entry.isDirectory()) add(join(dir, entry.name), `${prefix}${entry.name}/`);
      else parts.set(`${prefix}${entry.name}`, toBytes(readFileSync(join(dir, entry.name), 'utf8')));
    }
  };
  add(IDENTITIES, '');
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function open(bytes: Uint8Array, clientId: number): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

/** A session seeded from the parsed document rather than from the bytes. */
async function seeded(bytes: Uint8Array, clientId: number): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  documentToYrs(session, await parseDocx(bytes.slice().buffer, { preloadFonts: false }));
  return session;
}

/**
 * The saved document as `saveYrsDocx` and as the editor's Save write it, the
 * editor holding the source's comments and `added`.
 */
async function saves(
  session: YrsSession,
  added: Comment[] = []
): Promise<Array<[string, Uint8Array]>> {
  const yrs = (await saveYrsDocx(session)).bytes;
  const base = session.materializeDocx()!;
  const comments = [...(base.package.document.comments ?? []), ...added];
  const capture = captureSessionSave(session);
  const editor = await writeSessionSave(
    session,
    yrsToDocument(session, {
      ...base,
      package: { ...base.package, document: { ...base.package.document, comments } },
    }),
    capture,
    base.originalBuffer!,
    {},
    () => false
  );
  return [
    ['saveYrsDocx', yrs],
    ['editor', editor.bytes],
  ];
}

function documentXml(bytes: Uint8Array): string {
  return new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
}

/** The comment's markers, in document order. */
function markers(bytes: Uint8Array, id: number): string[] {
  return [
    ...documentXml(bytes).matchAll(
      new RegExp(`<w:comment(RangeStart|RangeEnd|Reference) w:id="${id}"/>`, 'g')
    ),
  ].map((match) => match[1]!);
}

function paragraphXml(bytes: Uint8Array, paraId: string): string {
  return documentXml(bytes).match(new RegExp(`<w:p [^>]*${paraId}.*?</w:p>`))![0];
}

/** The text a comment's first anchor covers. */
function anchored(session: YrsSession, commentId: string): string {
  const [anchor] = session.resolveComment(commentId);
  let offset = 0;
  let text = '';
  for (const segment of session.storySegments(anchor!.story)) {
    if (segment.kind === 'text') {
      text += segment.text.slice(
        Math.max(0, anchor!.start - offset),
        Math.max(0, anchor!.end - offset)
      );
      offset += segment.text.length;
    } else offset += 1;
  }
  return text;
}

function range(session: YrsSession, paraIndex: number, start: number, end: number) {
  const { paraId } = session.paragraphs('body')[paraIndex]!;
  return { story: 'body', start: { paraId, offset: start }, end: { paraId, offset: end } };
}

describe('a reanchored comment', () => {
  it('saves one range and one reference where it now is', async () => {
    const session = await open(docx(), 91001);
    session.beginUndoCapture();
    session.setCommentRanges('1', [range(session, 2, 0, 7)]);
    for (let round = 0; round < 2; round += 1) {
      for (const [path, bytes] of await saves(session)) {
        expect([path, markers(bytes, 1)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
        expect(paragraphXml(bytes, '0C000003')).toContain('<w:commentReference w:id="1"/>');
        expect(anchored(await open(bytes, 91002), '1')).toBe('Closing');
      }
      expect(session.undo()).toBe(true);
      for (const [path, bytes] of await saves(session)) {
        expect([path, markers(bytes, 1)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
        expect(anchored(await open(bytes, 91003), '1')).toBe('Intro');
      }
      expect(session.redo()).toBe(true);
    }
  });

  it('saves one reference once moved into a table cell', async () => {
    const session = await open(docx(), 91007);
    const cell = session.storyIds().find((story) => story !== 'body' && story.startsWith('body'))!;
    const { paraId } = session.paragraphs(cell)[0]!;
    session.setCommentRanges('1', [{ story: cell, start: { paraId, offset: 0 }, end: { paraId, offset: 4 } }]);
    for (const [path, bytes] of await saves(session)) {
      expect([path, markers(bytes, 1)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
      expect(paragraphXml(bytes, '0C000004')).toContain('<w:commentReference w:id="1"/>');
    }
  });

  it('saves one reference once replaced by a raw setComment', async () => {
    const session = await open(docx(), 91008);
    const start = session.locateParagraph('body', session.paragraphs('body')[2]!.paraId).start;
    session.applyRawOps('body', [{ op: 'setComment', id: '1', ranges: [[start + 8, start + 13]] }]);
    for (const [path, bytes] of await saves(session)) {
      expect([path, markers(bytes, 1)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
      expect(anchored(await open(bytes, 91009), '1')).toBe('words');
    }
  });

  it('saves one reference after replicas reanchor it concurrently, in either sync order', async () => {
    for (const peerFirst of [false, true]) {
      const source = await open(docx(), 91030);
      const peer = await createYrsSession({ clientId: 91031 });
      sessions.push(peer);
      peer.openDocx(docx(), false);
      peer.loadState(source.encodeState());
      source.setCommentRanges('1', [range(source, 2, 0, 7)]);
      peer.setCommentRanges('1', [range(peer, 3, 1, 6)]);
      const [first, second] = peerFirst ? [peer, source] : [source, peer];
      first.applyUpdate(second.encodeStateAsUpdate(first.encodeStateVector()));
      second.applyUpdate(first.encodeStateAsUpdate(second.encodeStateVector()));
      first.applyUpdate(second.encodeStateAsUpdate(first.encodeStateVector()));
      const expected = anchored(source, '1');
      expect(['Closing', 'After']).toContain(expected);
      const written = new Map<string, Set<string>>();
      for (const replica of [source, peer]) {
        expect(anchored(replica, '1')).toBe(expected);
        for (const [path, bytes] of await saves(replica)) {
          expect([path, markers(bytes, 1)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
          expect(anchored(await open(bytes, 91032), '1')).toBe(expected);
          written.set(path, (written.get(path) ?? new Set()).add(documentXml(bytes)));
        }
      }
      for (const [path, xml] of written) expect([path, xml.size]).toEqual([path, 1]);
    }
  });

  it('saves one reference when a start-only reanchor races a move, whichever wins', async () => {
    for (const peerFirst of [false, true]) {
      for (const sourceMoves of [false, true]) {
        const source = await open(docx(), 91040);
        const peer = await createYrsSession({ clientId: 91041 });
        sessions.push(peer);
        peer.openDocx(docx(), false);
        peer.loadState(source.encodeState());
        const [mover, keeper] = sourceMoves ? [source, peer] : [peer, source];
        keeper.setCommentRanges('1', [range(keeper, 0, 2, 5)]);
        mover.setCommentRanges('1', [range(mover, 2, 0, 7)]);
        const [first, second] = peerFirst ? [peer, source] : [source, peer];
        first.applyUpdate(second.encodeStateAsUpdate(first.encodeStateVector()));
        second.applyUpdate(first.encodeStateAsUpdate(second.encodeStateVector()));
        first.applyUpdate(second.encodeStateAsUpdate(first.encodeStateVector()));
        const expected = anchored(source, '1');
        expect(['tro', 'Closing']).toContain(expected);
        const written = new Map<string, Set<string>>();
        for (const replica of [source, peer]) {
          expect(anchored(replica, '1')).toBe(expected);
          for (const [path, bytes] of await saves(replica)) {
            expect([path, markers(bytes, 1)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
            expect(anchored(await open(bytes, 91042), '1')).toBe(expected);
            written.set(path, (written.get(path) ?? new Set()).add(documentXml(bytes)));
          }
        }
        for (const [path, xml] of written) expect([path, xml.size]).toEqual([path, 1]);
      }
    }
  });

  it('saves one reference in a replica that only receives a losing move', async () => {
    for (const newStory of [false, true]) {
      const keeper = await open(docx(), 91051);
      const mover = await createYrsSession({ clientId: 91050 });
      sessions.push(mover);
      mover.openDocx(docx(), false);
      mover.loadState(keeper.encodeState());
      keeper.setCommentRanges('1', [range(keeper, 0, 2, 5)]);
      if (newStory) {
        const at = { story: 'body', paraId: mover.paragraphs('body')[2]!.paraId, offset: 0 };
        const [cell] = mover.insertTable(at, 1, 1).createdStoryIds;
        const { paraId } = mover.paragraphs(cell!)[0]!;
        mover.insertText({ story: cell!, paraId, offset: 0 }, 'Fresh cell');
        mover.setCommentRanges('1', [
          { story: cell!, start: { paraId, offset: 0 }, end: { paraId, offset: 5 } },
        ]);
      } else {
        mover.setCommentRanges('1', [range(mover, 2, 0, 7)]);
      }
      keeper.applyUpdate(mover.encodeStateAsUpdate(keeper.encodeStateVector()));
      expect(anchored(keeper, '1')).toBe('tro');
      for (const [path, bytes] of await saves(keeper)) {
        expect([path, markers(bytes, 1)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
        expect(paragraphXml(bytes, '0C000001')).toContain('<w:commentReference w:id="1"/>');
        expect(anchored(await open(bytes, 91052), '1')).toBe('tro');
      }
    }
  });

  it('keeps every reference of a comment across stories when a replica loads it', async () => {
    const source = await open(identities(), 91060);
    const header = 'hf:rIdHeader';
    const lower = source.locateParagraph('body', '0000abcd').start;
    const { paraId: headerParaId } = source.paragraphs(header)[0]!;
    source.applyRawOps('body', [
      { op: 'setComment', id: '5', ranges: [[lower, lower + 2]], author: 'Ada', date: '2026-09-28T00:00:00Z' },
    ]);
    source.setCommentRanges('5', [
      { story: 'body', start: { paraId: '0000abcd', offset: 0 }, end: { paraId: '0000abcd', offset: 2 } },
      { story: header, start: { paraId: headerParaId, offset: 0 }, end: { paraId: headerParaId, offset: 4 } },
    ]);
    const reference = { modelKind: 'commentReference', commentId: 5 };
    source.applyRawOps('body', [{ op: 'insertEmbed', index: lower + 5, kind: 'field', payload: reference }]);
    const headerStart = source.locateParagraph(header, headerParaId).start;
    source.applyRawOps(header, [
      { op: 'insertEmbed', index: headerStart + 4, kind: 'field', payload: reference },
    ]);
    const replica = await createYrsSession({ clientId: 91061 });
    sessions.push(replica);
    replica.openDocx(identities(), false);
    replica.loadState(source.encodeState());
    for (const story of ['body', header]) {
      expect(replica.storySegments(story)).toEqual(source.storySegments(story));
    }
    for (const [path, bytes] of await saves(replica, [{ id: 5, author: 'Ada', content: [] }])) {
      expect([path, markers(bytes, 5)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
      const lowerXml = paragraphXml(bytes, '0000abcd');
      expect(lowerXml.indexOf('<w:commentReference w:id="5"/>')).toBeGreaterThan(lowerXml.indexOf('wer<'));
      const headerXml = new TextDecoder().decode(unzipContainer(bytes)['word/header1.xml']);
      expect(
        [...headerXml.matchAll(/<w:comment(RangeStart|RangeEnd|Reference) w:id="5"\/>/g)].map((match) => match[1])
      ).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
    }
  });

  it('saves once from a replica that received the reanchor', async () => {
    const source = await open(docx(), 91004);
    const peer = await createYrsSession({ clientId: 91005 });
    sessions.push(peer);
    peer.openDocx(docx(), false);
    peer.loadState(source.encodeState());
    source.setCommentRanges('1', [range(source, 2, 8, 13)]);
    peer.applyUpdate(source.encodeStateAsUpdate(peer.encodeStateVector()));
    for (const [path, bytes] of await saves(peer)) {
      expect([path, markers(bytes, 1)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
      expect(anchored(await open(bytes, 91006), '1')).toBe('words');
    }
  });
});

describe('a source comment', () => {
  for (const id of [0, 1]) {
    const bytes = () => docx(BODY.replaceAll('w:id="1"', `w:id="${id}"`), [id]);
    const reopened = async (saved: Uint8Array) => [await open(saved, 91071), await seeded(saved, 91072)];

    it(`with id ${id} reanchors with undo and redo and saves its new range`, async () => {
      const session = await open(bytes(), 91070);
      session.beginUndoCapture();
      session.setCommentRanges(String(id), [range(session, 2, 0, 7)]);
      expect(anchored(session, String(id))).toBe('Closing');
      expect(session.undo()).toBe(true);
      expect(anchored(session, String(id))).toBe('Intro');
      expect(session.redo()).toBe(true);
      for (const [path, saved] of await saves(session)) {
        expect([path, markers(saved, id)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
        for (const session of await reopened(saved)) expect(anchored(session, String(id))).toBe('Closing');
      }
    });

    it(`with id ${id} keeps its range after an edit elsewhere`, async () => {
      const session = await open(bytes(), 91073);
      const { paraId } = session.paragraphs('body')[2]!;
      session.insertText({ story: 'body', paraId, offset: 0 }, 'QA ');
      for (const [path, saved] of await saves(session)) {
        expect([path, markers(saved, id)]).toEqual([path, ['RangeStart', 'RangeEnd', 'Reference']]);
        for (const session of await reopened(saved)) expect(anchored(session, String(id))).toBe('Intro');
      }
    });
  }
});

describe('an added comment range', () => {
  const cases: Array<[string, number, number, number]> = [
    ['See the', 1, 0, 7],
    ['he l', 1, 5, 9],
    ['inked text then ins', 1, 9, 28],
    ['nse', 1, 26, 29],
    ['erted, ', 1, 28, 35],
    [' tail', 1, 37, 42],
    // The table before this paragraph is its first story unit.
    ['After', 3, 1, 6],
    ['to page', 4, 3, 10],
    ['age', 4, 7, 10],
    ['e  of', 4, 9, 15],
    ['of th', 4, 13, 18],
  ];
  for (const [expected, paragraph, start, end] of cases) {
    it(`saves paired markers and a reference around "${expected}"`, async () => {
      const session = await open(docx(), 91010);
      const { commentId } = session.addComment(
        [range(session, paragraph, start, end)],
        'Ada',
        '2026-09-28T00:00:00Z',
        'Note'
      );
      expect(anchored(session, commentId)).toBe(expected);
      const bytes = (await saveYrsDocx(session)).bytes;
      expect(markers(bytes, 2)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      expect(markers(bytes, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      for (const reopened of [await open(bytes, 91011), await seeded(bytes, 91012)]) {
        expect(anchored(reopened, '2')).toBe(expected);
        expect(reopened.paragraphs('body').map((paragraph) => paragraph.text)).toEqual(
          session.paragraphs('body').map((paragraph) => paragraph.text)
        );
      }
    });
  }
});

it('splits a hyperlink holding a field where a comment range ends inside it', async () => {
  const session = await open(docx(), 91040);
  session.setCommentRanges('1', [range(session, 4, 13, 18)]);
  const document = yrsToDocument(session, session.materializeDocx()!);
  const paragraph = document.package.document.content.find(
    (block) => block.type === 'paragraph' && block.paraId === '0C000006'
  ) as Paragraph;
  const text = (run: Run) =>
    run.content.map((entry) => (entry.type === 'text' ? entry.text : entry.type)).join('');
  expect(
    paragraph.content.map((child) =>
      child.type === 'hyperlink'
        ? (child.structuredChildren ?? child.children)
            .map((entry) => (entry.type === 'run' ? text(entry) : entry.type))
            .join('|')
        : child.type === 'run'
          ? text(child)
          : child.type
    )
  ).toEqual([
    'Go to ',
    'page |simpleField| ',
    'commentRangeStart',
    'of th',
    'commentRangeEnd',
    'commentReference',
    'e text',
    ' now',
  ]);
});

it('keeps a reference away from its range end and a field inside a hyperlink on a no-op save', async () => {
  const bytes = docx(
    `<w:p w14:paraId="0E000001"><w:commentRangeStart w:id="1"/>${t('Intro')}<w:commentRangeEnd w:id="1"/></w:p>` +
      `<w:p w14:paraId="0E000002"><w:r><w:commentReference w:id="1"/></w:r>${t('Next')}</w:p>` +
      `<w:p w14:paraId="0E000003"><w:hyperlink w:anchor="target">${t('Page ')}` +
      `<w:fldSimple w:instr=" PAGE ">${t('7')}</w:fldSimple></w:hyperlink></w:p>`
  );
  const saved = unzipContainer((await saveYrsDocx(await open(bytes, 91021))).bytes);
  expect(saved).toEqual(unzipContainer(bytes));
  expect(new TextDecoder().decode(saved['word/document.xml'])).toContain('<w:fldSimple w:instr=" PAGE ">');
});

it('keeps a commented document byte-identical on a no-op save', async () => {
  const bytes = docx();
  const saved = unzipContainer((await saveYrsDocx(await open(bytes, 91020))).bytes);
  expect(saved).toEqual(unzipContainer(bytes));
});

it('keeps the ranges of comments inside a table cell and a content control after an edit elsewhere', async () => {
  const commented = (id: number, text: string) =>
    `<w:commentRangeStart w:id="${id}"/>${t(text)}<w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:commentReference w:id="${id}"/></w:r>`;
  const body = BODY.replace(t('Cell'), commented(2, 'Cell')).replace(
    '<w:p w14:paraId="0C000005">',
    '<w:sdt><w:sdtPr><w:id w:val="6"/></w:sdtPr><w:sdtContent>' +
      `<w:p w14:paraId="0C000007">${commented(3, 'Controlled')}</w:p></w:sdtContent></w:sdt>` +
      '<w:p w14:paraId="0C000005">'
  );
  const session = await open(docx(body, [1, 2, 3]), 91080);
  const { paraId } = session.paragraphs('body')[2]!;
  session.insertText({ story: 'body', paraId, offset: 0 }, 'QA ');
  for (const [path, saved] of await saves(session)) {
    for (const [id, text] of [[1, 'Intro'], [2, 'Cell'], [3, 'Controlled']] as const) {
      expect([path, id, markers(saved, id)]).toEqual([path, id, ['RangeStart', 'RangeEnd', 'Reference']]);
      for (const reopened of [await open(saved, 91081), await seeded(saved, 91082)]) {
        expect(anchored(reopened, String(id))).toBe(text);
      }
    }
  }
});

it('keeps every range of a comment anchored in the body and a table cell across save and reopen', async () => {
  const session = await open(docx(), 91083);
  const cell = session.storyIds().find((story) => story !== 'body' && story.startsWith('body'))!;
  const { paraId } = session.paragraphs(cell)[0]!;
  session.setCommentRanges('1', [
    range(session, 2, 0, 7),
    { story: cell, start: { paraId, offset: 0 }, end: { paraId, offset: 4 } },
  ]);
  const ranges = session.resolveComment('1');
  expect(ranges.map(({ story }) => story)).toEqual(['body', cell]);
  let bytes = (await saveYrsDocx(session)).bytes;
  for (let cycle = 0; cycle < 2; cycle += 1) {
    const reopened = await open(bytes, 91084 + cycle);
    expect(reopened.resolveComment('1')).toEqual(ranges);
    const { paraId: after } = reopened.paragraphs('body')[3]!;
    reopened.insertText({ story: 'body', paraId: after, offset: 0 }, '!');
    bytes = (await saveYrsDocx(reopened)).bytes;
  }
});

it('keeps overlapping, enclosed, multi-paragraph and embed-only comment ranges after an unrelated edit', async () => {
  const start = (id: number) => `<w:commentRangeStart w:id="${id}"/>`;
  const end = (id: number) => `<w:commentRangeEnd w:id="${id}"/>`;
  const refs = (...ids: number[]) =>
    ids.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join('');
  const body =
    `<w:p w14:paraId="0F000001">${start(3)}${t('A')}${start(4)}${t('B')}${end(3)}${t('C')}` +
    `${start(6)}${t('D')}${end(6)}${end(4)}${refs(3, 4, 6)}</w:p>` +
    `<w:p w14:paraId="0F000002">${t('Before ')}${start(5)}${t('first')}</w:p>` +
    `<w:p w14:paraId="0F000003">${t('second')}${end(5)}${start(7)}` +
    `<w:fldSimple w:instr=" PAGE ">${t('1')}</w:fldSimple>${end(7)}${refs(5, 7)}${t(' after')}${start(8)}</w:p>` +
    '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
    `<w:p w14:paraId="0F00000A">${t('Cell')}</w:p></w:tc></w:tr></w:tbl>` +
    `<w:p w14:paraId="0F000004">${t('Past')}${end(8)}${refs(8)}${start(9)}</w:p>` +
    `<w:p w14:paraId="0F000005">${end(9)}${refs(9)}${t('Mark')}</w:p>` +
    `<w:p w14:paraId="0F000006">${t('Later')}</w:p>`;
  const ids = [3, 4, 6, 5, 7, 8, 9];
  const anchors = (session: YrsSession) => ids.map((id) => session.resolveComment(String(id)));
  const source = docx(body, ids);
  const session = await open(source, 91090);
  const expected = anchors(session);
  expect(ids.map((id) => anchored(session, String(id)))).toEqual(
    ['AB', 'BCD', 'D', 'firstsecond', '', 'Past', '']
  );
  expect(anchors(await seeded(source, 91093))).toEqual(expected);
  const later = session.paragraphs('body').at(-1)!;
  session.insertText({ story: 'body', paraId: later.paraId, offset: 0 }, 'QA ');
  for (const [path, bytes] of await saves(session)) {
    for (const id of ids) {
      expect([path, id, markers(bytes, id)]).toEqual([path, id, ['RangeStart', 'RangeEnd', 'Reference']]);
    }
    for (const reopened of [await open(bytes, 91091), await seeded(bytes, 91092)]) {
      expect([path, anchors(reopened)]).toEqual([path, expected]);
    }
    for (const [id, from, to] of [
      [5, '0F000002', '0F000003'],
      [8, '0F000003', '0F000004'],
      [9, '0F000004', '0F000005'],
    ] as const) {
      const spans = [paragraphXml(bytes, from).includes(start(id)), paragraphXml(bytes, to).includes(end(id))];
      expect([path, id, spans]).toEqual([path, id, [true, true]]);
    }
  }
});

it('keeps the body and table-cell ranges of a generated document after an edit to its title', async () => {
  const parts = unzipContainer(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '__fixtures__/comment-ranges/structure.docx')))
  );
  for (const name of ['word/document.xml', 'word/comments.xml']) {
    parts[name] = toBytes(new TextDecoder().decode(parts[name]).replaceAll('w:id="0"', 'w:id="2"'));
  }
  const session = await open(
    new Uint8Array(rezipPartsToArrayBuffer(new Map(Object.entries(parts)))),
    91100
  );
  const [title] = session.paragraphs('body');
  session.insertText({ story: 'body', paraId: title!.paraId, offset: 0 }, 'QA ');
  for (const [path, saved] of await saves(session)) {
    for (const [id, text] of [
      [2, 'Achado QA preservado. '],
      [1, 'Preservar comentário na célula'],
    ] as const) {
      expect([path, id, markers(saved, id)]).toEqual([path, id, ['RangeStart', 'RangeEnd', 'Reference']]);
      for (const reopened of [await open(saved, 91101), await seeded(saved, 91102)]) {
        expect(anchored(reopened, String(id))).toBe(text);
      }
    }
  }
});
