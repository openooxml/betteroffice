/**
 * The editor's own path into a note: a click resolves to a display position in
 * the note's story, the caret goes there, and typing at the caret has to reach
 * the saved file. `noteSave.test.ts` covers the projection with a caret the
 * test built by hand; this one builds it the way a pointer does.
 */

import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { DisplayList, DisplayListRegionHit, TextRunPrimitive } from '../layout/render';
import type { Paragraph } from '../types/content';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { createYrsInputPositionMap, displayPositionToYrsLoc } from './inputPositionMap';
import { captureSessionSave, saveYrsDocx, writeSessionSave } from './saveYrsDocx';
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
</Relationships>`;

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Body text</w:t></w:r>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>
      <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="2"/></w:r>
    </w:p>
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

const FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="2">
    <w:p>
      <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
      <w:r><w:t xml:space="preserve">Ibsen 1879</w:t></w:r>
    </w:p>
  </w:footnote>
  <w:footnote w:id="3"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t>Clean footnote sibling.</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;

const ENDNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:id="2">
    <w:p>
      <w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>
      <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>
      <w:r><w:t xml:space="preserve">Ibsen 1879</w:t></w:r>
    </w:p>
  </w:endnote>
  <w:endnote w:id="3"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t>Clean endnote sibling.</w:t></w:r></w:p></w:endnote>
</w:endnotes>`;

function fixture(overrides: PartsMap = new Map()): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/_rels/document.xml.rels', toBytes(DOCUMENT_RELS));
  parts.set('word/document.xml', toBytes(DOCUMENT));
  parts.set('word/styles.xml', toBytes(STYLES));
  parts.set('word/footnotes.xml', toBytes(FOOTNOTES));
  parts.set('word/endnotes.xml', toBytes(ENDNOTES));
  for (const [part, bytes] of overrides) parts.set(part, bytes);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/**
 * Place the caret from a display position in `story` — the hit test's own
 * output — and type there, the way the pointer path and YrsInput do.
 */
function typeAtDisplayPosition(
  session: YrsSession,
  story: string,
  position: number,
  text: string
): void {
  const map = createYrsInputPositionMap(story, session.paragraphSpans(story));
  const caret = displayPositionToYrsLoc(map, position);
  if (!caret) throw new Error(`no location at display position ${position} of ${story}`);
  session.setSelection(caret);
  const head = session.selection()?.head;
  if (!head) throw new Error(`the selection did not land in ${story}`);
  session.insertText(head, text);
}

interface NoteClick {
  story: string;
  position: number;
}

function clickedNotePositions(session: YrsSession): NoteClick[] {
  const layout = session.layoutDocumentWithRegionsJson(
    JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {} }] },
      notes: {
        contents: [
          { id: 2, noteKind: 'footnote', height: 0 },
          { id: 2, noteKind: 'endnote', height: 0 },
        ],
      },
      measurement: {
        fontChains: {},
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: false,
      },
      renderEnv: {},
    })
  );
  const list = JSON.parse(session.buildDisplayListJson(layout)) as DisplayList;
  return (['footnote', 'endnote'] as const).map((kind) => {
    const located = list.pages
      .flatMap((page) => (page.noteAreas ?? []).map((area) => ({ page, area })))
      .find(({ area }) => (area.kind ?? 'footnote') === kind && (area.noteIds ?? []).includes(2));
    if (!located) throw new Error(`no ${kind} area`);
    const run = located.area.primitives?.find(
      (primitive): primitive is TextRunPrimitive & { docStart: number } =>
        primitive.kind === 'text' && primitive.docStart != null && primitive.text.includes('Ibsen')
    );
    if (!run) throw new Error(`no positioned ${kind} text`);
    const split = run.text.indexOf(' ');
    if (split < 0) throw new Error(`no word boundary in ${kind} text`);
    const target = run.docStart + split;
    for (let step = 0; step <= 512; step += 1) {
      const x = run.x + (run.width * step) / 512;
      const hit = JSON.parse(
        session.displayHitTestRegionsJson(located.page.pageIndex, x, run.baselineY - 2)
      ) as DisplayListRegionHit | null;
      if (hit?.region !== kind || hit.noteId !== 2 || hit.pos !== target) continue;
      return {
        story: `${kind === 'footnote' ? 'fn' : 'en'}:${hit.noteId}`,
        position: hit.pos,
      };
    }
    throw new Error(`no ${kind} hit resolved display position ${target}`);
  });
}

function noteXml(xml: string, tag: 'footnote' | 'endnote', id: number): string {
  const match = new RegExp(
    `<w:${tag}\\b[^>]*\\bw:id=["']${id}["'][^>]*>[\\s\\S]*?</w:${tag}>`
  ).exec(xml);
  if (!match) throw new Error(`no ${tag} ${id}`);
  return match[0];
}

function rawNoteText(xml: string): string {
  return [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join('');
}

describe('typing into a note through the editor path', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  for (const kind of ['footnote', 'endnote'] as const) {
    it(`keeps an untouched ${kind} number in a mixed text run`, async () => {
      const part = `word/${kind}s.xml`;
      const source = kind === 'footnote' ? FOOTNOTES : ENDNOTES;
      const sibling = noteXml(source, kind, 3);
      const notes = source.replace(
        sibling,
        `<w:${kind} w:id="3"><w:p w:rsidR="00AB12CD"><w:r><w:${kind}Ref/><w:t>Text</w:t></w:r></w:p></w:${kind}>`
      );
      const bytes = fixture(new Map([[part, toBytes(notes)]]));
      const session = await createYrsSession({ clientId: 53007 });

      try {
        session.seedFromDocx(bytes);
        const story = `${kind === 'footnote' ? 'fn' : 'en'}:2`;
        typeAtDisplayPosition(session, story, 0, 'Revised ');

        const saved = await saveYrsDocx(session);
        const xml = readDocxContainer(saved.bytes.slice().buffer).text(part) ?? '';
        expect(rawNoteText(noteXml(xml, kind, 2))).toBe('Revised Ibsen 1879');
        expect(noteXml(xml, kind, 3)).toBe(noteXml(notes, kind, 3));
      } finally {
        session.destroy();
      }
    });

    it(`keeps the ${kind} number when accepting deletion of its first paragraph`, async () => {
      const part = `word/${kind}s.xml`;
      const source = kind === 'footnote' ? FOOTNOTES : ENDNOTES;
      const opening = noteXml(source, kind, 2);
      const notes = source.replace(
        opening,
        opening.replace(
          `</w:${kind}>`,
          `<w:p><w:r><w:t>Kept second paragraph.</w:t></w:r></w:p></w:${kind}>`
        )
      );
      const bytes = fixture(new Map([[part, toBytes(notes)]]));
      const session = await createYrsSession({ clientId: 53006 });

      try {
        session.seedFromDocx(bytes);
        const story = `${kind === 'footnote' ? 'fn' : 'en'}:2`;
        const [first, second] = session.paragraphs(story);
        const { revisionId } = session.deleteRange(
          {
            story,
            start: { paraId: first.paraId, offset: 0 },
            end: { paraId: second.paraId, offset: 0 },
          },
          { name: 'Reviewer', date: '2026-10-21T00:00:00Z' }
        );
        expect(revisionId).not.toBeNull();
        session.acceptChange({ revisionId: revisionId! });
        expect(session.paragraphs(story).map(({ paraId, text }) => ({ paraId, text }))).toEqual([
          { paraId: second.paraId, text: 'Kept second paragraph.' },
        ]);

        const splice = session.paragraphSavePlan().splicedParts?.find(
          (entry) => entry.part === part
        );
        expect(splice).toBeDefined();
        expect(splice!.removed).toContain(2);
        expect(splice!.paragraphs).toContainEqual([3, second.paraId]);
        expect(splice!.changed).not.toContain(3);

        const saved = await saveYrsDocx(session);
        const parts = readDocxContainer(saved.bytes.slice().buffer);
        const xml = parts.text(part) ?? '';
        const edited = noteXml(xml, kind, 2);
        const paragraphs = [...edited.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)];
        const mark = new RegExp(`<w:${kind}Ref\\b`, 'g');
        expect(paragraphs).toHaveLength(1);
        expect([...edited.matchAll(mark)]).toHaveLength(1);
        expect([...paragraphs[0][0].matchAll(mark)]).toHaveLength(1);
        expect(rawNoteText(edited)).toBe('Kept second paragraph.');
        for (const id of [-1, 0, 3]) {
          expect(noteXml(xml, kind, id)).toBe(noteXml(notes, kind, id));
        }
        const otherKind = kind === 'footnote' ? 'endnote' : 'footnote';
        const otherSource = kind === 'footnote' ? ENDNOTES : FOOTNOTES;
        const otherXml = parts.text(`word/${otherKind}s.xml`) ?? '';
        for (const id of [-1, 0, 2, 3]) {
          expect(noteXml(otherXml, otherKind, id)).toBe(noteXml(otherSource, otherKind, id));
        }
      } finally {
        session.destroy();
      }
    });
  }

  for (const kind of ['footnote', 'endnote'] as const) {
    it(`writes ${kind} separator edits saved alongside a spliced note edit`, async () => {
      const part = `word/${kind}s.xml`;
      const otherPart = kind === 'footnote' ? 'word/endnotes.xml' : 'word/footnotes.xml';
      const bytes = fixture();
      const session = await createYrsSession({ clientId: 53009 });

      try {
        session.seedFromDocx(bytes);
        typeAtDisplayPosition(session, `${kind === 'footnote' ? 'fn' : 'en'}:2`, 0, 'Revised ');
        const capture = captureSessionSave(session);
        const entry = capture.plan.splicedParts?.find((candidate) => candidate.part === part);
        expect(entry?.paragraphPatches?.length).toBeGreaterThan(0);

        const base = session.materializeDocx()!;
        const document = yrsToDocument(session, base);
        const key = kind === 'footnote' ? 'footnoteSeparators' : 'endnoteSeparators';
        const [separator, ...rest] = document.package[key]!;
        const [paragraph] = separator.content as Paragraph[];
        const edited: Paragraph = {
          ...paragraph,
          content: [{ type: 'run', content: [{ type: 'text', text: 'Edited separator' }] }],
        };
        const output: Document = {
          ...document,
          package: { ...document.package, [key]: [{ ...separator, content: [edited] }, ...rest] },
        };

        const whole = await writeSessionSave(
          session,
          output,
          capture,
          base.originalBuffer!,
          {},
          () => false
        );
        const saved = await writeSessionSave(session, output, capture, base.originalBuffer!);
        const parts = readDocxContainer(saved.bytes.slice().buffer);
        const xml = parts.text(part) ?? '';
        expect(xml).toBe(readDocxContainer(whole.bytes.slice().buffer).text(part) ?? '');
        expect(rawNoteText(noteXml(xml, kind, -1))).toBe('Edited separator');
        expect(rawNoteText(noteXml(xml, kind, 2))).toBe('Revised Ibsen 1879');
        expect(parts.text(otherPart)).toBe(readDocxContainer(bytes.slice().buffer).text(otherPart));
      } finally {
        session.destroy();
      }
    });
  }

  for (const kind of ['footnote', 'endnote'] as const) {
    for (const operation of ['split', 'delete'] as const) {
      it(`keeps a table-bearing ${kind} part when paragraphs ${operation}`, async () => {
        const part = `word/${kind}s.xml`;
        const source = kind === 'footnote' ? FOOTNOTES : ENDNOTES;
        const opening = noteXml(source, kind, 2);
        const sibling = noteXml(source, kind, 3);
        const notes = source.replace(
          opening,
          opening.replace(
            `</w:${kind}>`,
            `<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p></w:${kind}>`
          )
        ).replace(
          sibling,
          sibling.replace(
            `</w:${kind}>`,
            '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="none" w:sz="4"/></w:tblBorders></w:tblPr>' +
              '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr w:rsidR="00AB12CD">' +
              `<w:tc><w:p><w:r><w:t>Table text.</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:${kind}>`
          )
        );
        const bytes = fixture(new Map([[part, toBytes(notes)]]));
        const session = await createYrsSession({ clientId: 53008 });

        try {
          session.seedFromDocx(bytes);
          const story = `${kind === 'footnote' ? 'fn' : 'en'}:2`;
          const [first, second] = session.paragraphs(story);
          if (operation === 'split') {
            session.splitParagraph({ story, paraId: first.paraId, offset: 5 });
          } else {
            session.deleteRange({
              story,
              start: { paraId: first.paraId, offset: first.text.length },
              end: { paraId: second.paraId, offset: second.text.length },
            });
          }

          const splice = session.paragraphSavePlan().splicedParts?.find(
            (entry) => entry.part === part
          );
          expect(splice).toBeDefined();
          expect(operation === 'split' ? splice!.inserted : splice!.removed).toHaveLength(1);

          const saved = await saveYrsDocx(session);
          const xml = readDocxContainer(saved.bytes.slice().buffer).text(part) ?? '';
          const edited = noteXml(xml, kind, 2);
          expect(rawNoteText(edited)).toBe(
            operation === 'split' ? 'Ibsen 1879Second paragraph.' : 'Ibsen 1879'
          );
          expect([...edited.matchAll(/<w:p(?:\s[^>]*)?>/g)]).toHaveLength(
            operation === 'split' ? 3 : 1
          );
          for (const id of [-1, 0, 3]) {
            expect(noteXml(xml, kind, id)).toBe(noteXml(notes, kind, id));
          }
        } finally {
          session.destroy();
        }
      });
    }

    it(`keeps the other ${kind}s' source XML when replacing a word in one ${kind}`, async () => {
      const part = `word/${kind}s.xml`;
      const story = `${kind === 'footnote' ? 'fn' : 'en'}:2`;
      const source = kind === 'footnote' ? FOOTNOTES : ENDNOTES;
      const notes = source.replace(
        `<w:${kind}s xmlns:w=`,
        `<w:${kind}s xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w=`
      ).replace(
        `<w:r><w:t>Clean ${kind} sibling.</w:t></w:r>`,
        `<w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t xml:space="preserve">Clean ${kind} </w:t></w:r>` +
          '<w:hyperlink r:id="rIdLink" w:history="1"><w:r><w:t>sibling.</w:t></w:r></w:hyperlink>'
      ).replace(
        `</w:${kind}s>`,
        `<w:${kind} w:id="4" w:type="continuationNotice"><w:p><w:r><w:t>Continued.</w:t></w:r></w:p></w:${kind}></w:${kind}s>`
      );
      const bytes = fixture(new Map([
        [part, toBytes(notes)],
        [`word/_rels/${kind}s.xml.rels`, toBytes(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/note" TargetMode="External"/>
</Relationships>`)],
      ]));
      const session = await createYrsSession({ clientId: 53005 });

      try {
        session.seedFromDocx(bytes);
        const { paraId } = session.paragraphs(story)[0];
        session.deleteRange({
          story,
          start: { paraId, offset: 0 },
          end: { paraId, offset: 5 },
        });
        typeAtDisplayPosition(session, story, 0, 'Revised');

        const splice = session.paragraphSavePlan().splicedParts?.find(
          (entry) => entry.part === part
        );
        expect(splice).toBeDefined();
        expect(splice!.changed).toHaveLength(1);
        expect(splice!.paragraphs).toContainEqual([2, paraId]);
        expect(splice!.paragraphs.map(([ordinal]) => ordinal)).toEqual([0, 1, 2, 3, 4]);
        for (const ordinal of [0, 1, 4]) expect(splice!.changed).not.toContain(ordinal);

        const saved = await saveYrsDocx(session);
        const parts = readDocxContainer(saved.bytes.slice().buffer);
        const xml = parts.text(part) ?? '';
        const edited = noteXml(xml, kind, 2);
        expect(rawNoteText(edited)).toBe('Revised 1879');
        expect(edited).toContain(`<w:${kind}Ref/>`);
        const style = kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference';
        expect(edited).toContain(`<w:rStyle w:val="${style}"/>`);
        for (const id of [-1, 0, 3, 4]) {
          expect(noteXml(xml, kind, id)).toBe(noteXml(notes, kind, id));
        }
      } finally {
        session.destroy();
      }
    });
  }

  it('selectively saves note clicks without rewriting clean stories', async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 53003 });

    let saved: Document;
    try {
      session.seedFromDocx(bytes);
      const clicks = clickedNotePositions(session);
      expect(clicks.map((click) => click.story)).toEqual(['fn:2', 'en:2']);
      for (const click of clicks) typeAtDisplayPosition(session, click.story, click.position, ',');
      saved = yrsToDocument(session, parsed, {
        storyIds: new Set(clicks.map((click) => click.story)),
      });
    } finally {
      session.destroy();
    }

    expect(saved.package.document.content).toBe(parsed.package.document.content);

    const parts = readDocxContainer(await repackDocx(saved));
    const footnotesXml = parts.text('word/footnotes.xml') ?? '';
    const endnotesXml = parts.text('word/endnotes.xml') ?? '';
    const footnote = noteXml(footnotesXml, 'footnote', 2);
    const endnote = noteXml(endnotesXml, 'endnote', 2);

    expect(footnote).toContain('<w:footnoteRef/>');
    expect(rawNoteText(footnote)).toBe('Ibsen, 1879');
    expect(endnote).toContain('<w:endnoteRef/>');
    expect(rawNoteText(endnote)).toBe('Ibsen, 1879');
    expect(noteXml(footnotesXml, 'footnote', 3)).toBe(noteXml(FOOTNOTES, 'footnote', 3));
    expect(noteXml(endnotesXml, 'endnote', 3)).toBe(noteXml(ENDNOTES, 'endnote', 3));
  });

  it('selectively saves a body undo while the note caret is active', async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 53004 });

    try {
      session.seedFromDocx(bytes);
      const body = session.paragraphs('body')[0];
      session.setSelection({ story: 'body', paraId: body.paraId, offset: 9 });
      session.insertText(session.selection()!.head, '!');
      const afterBodyEdit = yrsToDocument(session, parsed, {
        storyIds: new Set(['body']),
      });
      const note = session.paragraphs('fn:2')[0];
      session.setSelection({ story: 'fn:2', paraId: note.paraId, offset: 0 });

      expect(session.undo()).toBe(true);
      expect(session.paragraphs('body')[0].text).toBe('Body text');
      expect(session.historyStories()).toEqual(['body']);

      const misattributed = yrsToDocument(session, afterBodyEdit, {
        storyIds: new Set(['fn:2']),
      });
      const misattributedParts = readDocxContainer(await repackDocx(misattributed));
      expect(rawNoteText(misattributedParts.text('word/document.xml') ?? '')).toBe('Body text!');

      const saved = yrsToDocument(session, afterBodyEdit, {
        storyIds: new Set(session.historyStories()),
      });
      const savedParts = readDocxContainer(await repackDocx(saved));
      expect(rawNoteText(savedParts.text('word/document.xml') ?? '')).toBe('Body text');
    } finally {
      session.destroy();
    }
  });
});
