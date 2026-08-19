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
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { createYrsInputPositionMap, displayPositionToYrsLoc } from './inputPositionMap';
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

function fixture(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/_rels/document.xml.rels', toBytes(DOCUMENT_RELS));
  parts.set('word/document.xml', toBytes(DOCUMENT));
  parts.set('word/styles.xml', toBytes(STYLES));
  parts.set('word/footnotes.xml', toBytes(FOOTNOTES));
  parts.set('word/endnotes.xml', toBytes(ENDNOTES));
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
      expect(session.historyStory()).toBe('body');
      const afterBodyEdit = yrsToDocument(session, parsed, {
        storyIds: new Set(['body']),
      });
      const note = session.paragraphs('fn:2')[0];
      session.setSelection({ story: 'fn:2', paraId: note.paraId, offset: 0 });

      expect(session.historyStory()).toBe('body');
      expect(session.undo()).toBe(true);
      expect(session.paragraphs('body')[0].text).toBe('Body text');

      const misattributed = yrsToDocument(session, afterBodyEdit, {
        storyIds: new Set(['fn:2']),
      });
      const misattributedParts = readDocxContainer(await repackDocx(misattributed));
      expect(rawNoteText(misattributedParts.text('word/document.xml') ?? '')).toBe('Body text!');

      const historyStory = session.historyStory();
      expect(historyStory).toBe('body');
      const saved = yrsToDocument(session, afterBodyEdit, {
        storyIds: new Set([historyStory!]),
      });
      const savedParts = readDocxContainer(await repackDocx(saved));
      expect(rawNoteText(savedParts.text('word/document.xml') ?? '')).toBe('Body text');
    } finally {
      session.destroy();
    }
  });
});
