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
import type { BlockContent, Document, Endnote, Footnote } from '../types/document';
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
      <w:r><w:t xml:space="preserve">Ibsen 1879</w:t></w:r>
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
      <w:r><w:t xml:space="preserve">Ibsen 1879</w:t></w:r>
    </w:p>
  </w:endnote>
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

describe('typing into a note through the editor path', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('saves what was typed where the caret was clicked', async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 53003 });

    let saved: Document;
    try {
      session.seedFromDocx(bytes);
      // display position 6 = five characters into "Ibsen 1879", i.e. between
      // the name and the year — a caret a click could put there
      typeAtDisplayPosition(session, 'fn:2', 6, ',');
      typeAtDisplayPosition(session, 'en:2', 6, ',');
      saved = yrsToDocument(session, parsed);
    } finally {
      session.destroy();
    }

    const reopened = await parseDocx(await repackDocx(saved), { preloadFonts: false });
    expect(noteText(reopened.package.footnotes, 2)).toBe('Ibsen, 1879');
    expect(noteText(reopened.package.endnotes, 2)).toBe('Ibsen, 1879');
  });
});
