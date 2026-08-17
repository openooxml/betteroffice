import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { BlockContent, Document, Endnote, Footnote } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${OFFICE_DOC}.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="${OFFICE_DOC}.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="${OFFICE_DOC}.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/comments.xml" ContentType="${OFFICE_DOC}.wordprocessingml.comments+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;

const BODY = `<w:p>
      <w:r><w:t xml:space="preserve">Body text</w:t></w:r>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>
      <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="2"/></w:r>
    </w:p>`;

/** `A`, a reference to note `id`, `B`, with a bookmark closing right after the reference. */
const bookmarkedBody = (id: number, trailing: string): string => `<w:p>
      <w:r><w:t>A</w:t></w:r>
      <w:bookmarkStart w:id="5" w:name="span"/>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>
      <w:bookmarkEnd w:id="5"/>
      ${trailing}
    </w:p>`;

const documentXml = (body: string): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
</w:styles>`;

const FOOTNOTE_BODY = `<w:r><w:t xml:space="preserve"> Footnote body.</w:t></w:r>`;

const COMMENTED_FOOTNOTE_BODY = `<w:r><w:t xml:space="preserve"> Footnote </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>body</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:t>.</w:t></w:r>`;

const REVISED_FOOTNOTE_BODY = `<w:r>
        <w:rPr><w:rPrChange w:id="7" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr>
        <w:t xml:space="preserve"> Footnote body.</w:t>
      </w:r>`;

const MARK_RUN = `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>`;

const DOUBLE_MARK_RUN = `<w:r>
        <w:rPr><w:rStyle w:val="FootnoteReference"/><w:rPrChange w:id="9" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr>
        <w:footnoteRef/><w:footnoteRef/>
      </w:r>`;

const REVISED_MARK_RUN = `<w:r>
        <w:rPr><w:rStyle w:val="FootnoteReference"/><w:rPrChange w:id="8" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>
        <w:footnoteRef/>
      </w:r>`;

const DELETED_MARK_RUN = `<w:del w:id="6" w:author="Reviewer" w:date="2026-08-16T00:00:00Z">
        <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
      </w:del>`;

const footnotesXml = (body: string, markRun: string, noteId: number): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="${noteId}">
    <w:p>
      <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
      ${markRun}
      ${body}
    </w:p>
  </w:footnote>
</w:footnotes>`;

const ENDNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:id="2">
    <w:p>
      <w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>
      <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>
      <w:r><w:t xml:space="preserve"> Endnote body.</w:t></w:r>
    </w:p>
  </w:endnote>
</w:endnotes>`;

const COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"><w:p><w:r><w:t>Note comment</w:t></w:r></w:p></w:comment>
</w:comments>`;

interface FixtureOptions {
  footnoteBody?: string;
  markRun?: string;
  body?: string;
  noteId?: number;
}

function fixture(options: FixtureOptions = {}): Uint8Array {
  const {
    footnoteBody = FOOTNOTE_BODY,
    markRun = MARK_RUN,
    body = BODY,
    noteId = 2,
  } = options;
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/_rels/document.xml.rels', toBytes(DOCUMENT_RELS));
  parts.set('word/document.xml', toBytes(documentXml(body)));
  parts.set('word/styles.xml', toBytes(STYLES));
  parts.set('word/footnotes.xml', toBytes(footnotesXml(footnoteBody, markRun, noteId)));
  parts.set('word/endnotes.xml', toBytes(ENDNOTES));
  parts.set('word/comments.xml', toBytes(COMMENTS));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

interface SaveOptions {
  seeder?: 'native' | 'projected';
  noteId?: number;
  edit?: (session: YrsSession) => void;
}

interface SaveResult {
  /** Inner XML of the saved footnote. */
  xml: string;
  /** Saved `word/document.xml`. */
  body: string;
  /** What a peer sees on the wire for the note story. */
  segments: ReturnType<YrsSession['storySegments']>;
  reopened: Document;
}

/** Seed `bytes`, run the edit, and repack. */
async function saveFootnote(bytes: Uint8Array, options: SaveOptions = {}): Promise<SaveResult> {
  const { seeder = 'native', noteId = 2, edit } = options;
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const session = await createYrsSession({ clientId: 53003 });
  let saved: Document;
  let segments: SaveResult['segments'];
  try {
    if (seeder === 'native') session.seedFromDocx(bytes);
    else documentToYrs(session, parsed);
    edit?.(session);
    segments = session.storySegments(`fn:${noteId}`);
    saved = yrsToDocument(session, parsed);
  } finally {
    session.destroy();
  }
  const repacked = await repackDocx(saved);
  const parts = readDocxContainer(repacked);
  return {
    xml: noteXml(parts.text('word/footnotes.xml') ?? '', 'footnote', noteId),
    body: parts.text('word/document.xml') ?? '',
    segments,
    reopened: await parseDocx(repacked, { preloadFonts: false }),
  };
}

/** Type at the end of a note story, the way the editor would. */
function appendToNote(session: YrsSession, story: string, text: string): void {
  const spans = session.paragraphSpans(story);
  const last = spans[spans.length - 1];
  session.insertText({ story, paraId: last.paraId, offset: last.length }, text);
}

function blocksText(blocks: readonly BlockContent[]): string {
  return blocks
    .map((block) =>
      block.type === 'paragraph'
        ? block.content
            .map((item) =>
              item.type === 'run'
                ? item.content.map((run) => (run.type === 'text' ? run.text : '')).join('')
                : ''
            )
            .join('')
        : ''
    )
    .join('');
}

function noteText(notes: readonly (Footnote | Endnote)[] | undefined, id: number): string {
  const note = notes?.find((candidate) => candidate.id === id);
  return note ? blocksText(note.content) : '';
}

function noteXml(xml: string, tag: 'footnote' | 'endnote', id: number): string {
  const match = new RegExp(`<w:${tag} w:id="${id}"[^>]*>([\\s\\S]*?)</w:${tag}>`).exec(xml);
  return match?.[1] ?? '';
}

describe('note story save projection', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('round-trips text typed into a footnote and into an endnote', async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 53002 });

    let saved: Document;
    try {
      session.seedFromDocx(bytes);
      appendToNote(session, 'fn:2', ' Typed into the footnote.');
      appendToNote(session, 'en:2', ' Typed into the endnote.');
      saved = yrsToDocument(session, parsed);
    } finally {
      session.destroy();
    }

    const repacked = await repackDocx(saved);
    const reopened = await parseDocx(repacked, { preloadFonts: false });
    expect(noteText(reopened.package.footnotes, 2)).toContain('Typed into the footnote.');
    expect(noteText(reopened.package.endnotes, 2)).toContain('Typed into the endnote.');

    const parts = readDocxContainer(repacked);
    const footnote = noteXml(parts.text('word/footnotes.xml') ?? '', 'footnote', 2);
    expect(footnote).toContain('<w:footnoteRef');
    expect(footnote).toMatch(/<w:rStyle w:val="FootnoteReference"\/>[\s\S]*?<w:footnoteRef/);
    expect(footnote).toContain(' Footnote body. Typed into the footnote.');
    const endnote = noteXml(parts.text('word/endnotes.xml') ?? '', 'endnote', 2);
    expect(endnote).toContain('<w:endnoteRef');
    expect(endnote).toMatch(/<w:rStyle w:val="EndnoteReference"\/>[\s\S]*?<w:endnoteRef/);
    expect(endnote).toContain(' Endnote body. Typed into the endnote.');
  });

  it('keeps the note mark when the reviewer suggests deleting the note text', async () => {
    const { xml, reopened } = await saveFootnote(fixture(), {
      edit: (session) => {
        const [{ paraId, length }] = session.paragraphSpans('fn:2');
        session.deleteRange(
          { story: 'fn:2', start: { paraId, offset: 0 }, end: { paraId, offset: length } },
          { name: 'Reviewer', date: '2026-08-16T00:00:00Z' }
        );
      },
    });

    expect(xml).toContain('<w:footnoteRef/>');
    expect(xml).toMatch(/<w:del [^>]*>[\s\S]*?Footnote body\.[\s\S]*?<\/w:del>/);
    const paragraph = reopened.package.footnotes?.find((note) => note.id === 2)?.content[0];
    const mark = paragraph?.type === 'paragraph' ? paragraph.content[0] : undefined;
    expect(
      mark?.type === 'run' && mark.content.some((entry) => entry.type === 'footnoteRefMark')
    ).toBe(true);
  });

  it('keeps a note mark the source file already had inside a suggested deletion', async () => {
    const { xml } = await saveFootnote(fixture({ markRun: DELETED_MARK_RUN }), {
      edit: (session) => appendToNote(session, 'fn:2', ' Typed.'),
    });

    expect(xml).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>)[\s\S])*<w:footnoteRef\/>[\s\S]*?<\/w:del>/);
    expect(xml).toContain(' Footnote body. Typed.');
  });

  it('places note comment ranges after the mark on the right characters', async () => {
    const { xml } = await saveFootnote(fixture({ footnoteBody: COMMENTED_FOOTNOTE_BODY }), {
      edit: (session) => appendToNote(session, 'fn:2', ' Typed.'),
    });

    expect(xml).toContain('<w:footnoteRef');
    expect(xml).toMatch(
      /<w:commentRangeStart w:id="1"\/><w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t>body<\/w:t><\/w:r><w:commentRangeEnd w:id="1"\/>/
    );
  });

  it('restores original note runs with their property changes when the text is untouched', async () => {
    const { xml } = await saveFootnote(fixture({ footnoteBody: REVISED_FOOTNOTE_BODY }));

    expect(xml).toContain('<w:footnoteRef');
    expect(xml).toMatch(/<w:rPrChange [^>]*w:id="7"/);
    expect(xml).toContain(' Footnote body.');
  });

  for (const seeder of ['native', 'projected'] as const) {
    it(`[${seeder}] leaves the note number mark off the wire`, async () => {
      const { segments, xml } = await saveFootnote(fixture(), { seeder });

      expect(segments.filter((segment) => segment.kind === 'embed')).toEqual([]);
      expect(segments[0]).toMatchObject({ kind: 'text', text: ' Footnote body.' });
      expect(xml).toContain('<w:footnoteRef/>');
    });

    it(`[${seeder}] keeps both marks of a run and the run's property change`, async () => {
      const { xml } = await saveFootnote(fixture({ markRun: DOUBLE_MARK_RUN }), { seeder });

      expect(xml.match(/<w:footnoteRef\/>/g)?.length).toBe(2);
      expect(xml).toMatch(/<w:rPrChange [^>]*w:id="9"/);
      expect(xml).toContain(' Footnote body.');
    });

    it(`[${seeder}] keeps a mark run's property change on that run once the text is edited`, async () => {
      const { xml } = await saveFootnote(fixture({ markRun: REVISED_MARK_RUN }), {
        seeder,
        edit: (session) => appendToNote(session, 'fn:2', ' Typed.'),
      });

      expect(xml).toContain(' Footnote body. Typed.');
      const runs = xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
      expect(runs.filter((run) => /<w:rPrChange [^>]*w:id="8"/.test(run))).toEqual([
        expect.stringContaining('<w:footnoteRef/>'),
      ]);
    });

    it(`[${seeder}] closes a bookmark right after a multi-digit note reference`, async () => {
      const { body } = await saveFootnote(
        fixture({ body: bookmarkedBody(12, '<w:r><w:t>B</w:t></w:r>'), noteId: 12 }),
        { seeder, noteId: 12 }
      );

      expect(body).toMatch(
        /<w:footnoteReference w:id="12"\/><\/w:r><w:bookmarkEnd w:id="5"\/><w:r><w:t>B<\/w:t>/
      );
    });

    it(`[${seeder}] keeps a bookmark that closes on a trailing note reference`, async () => {
      const { body } = await saveFootnote(
        fixture({ body: bookmarkedBody(12, ''), noteId: 12 }),
        { seeder, noteId: 12 }
      );

      expect(body).toContain('<w:bookmarkEnd w:id="5"/>');
    });
  }
});
