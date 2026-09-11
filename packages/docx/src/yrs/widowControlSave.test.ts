import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Document, Paragraph } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${OFFICE_DOC}.wordprocessingml.styles+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Authored after import</w:t></w:r></w:p>
    <w:p><w:pPr><w:widowControl w:val="0"/></w:pPr><w:r><w:t>Cleared after import</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

function fixture(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/document.xml', toBytes(DOCUMENT));
  parts.set('word/styles.xml', toBytes(STYLES));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function bodyParagraphs(document: Document): Paragraph[] {
  return document.package.document.content.filter(
    (block): block is Paragraph => block.type === 'paragraph'
  );
}

describe('widow control save projection', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('round-trips a newly authored false and a cleared original false', async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, {
      preloadFonts: false,
    });
    const session = await createYrsSession({ clientId: 53001 });

    let saved: Document;
    try {
      session.seedFromDocx(bytes);
      const [authored, cleared] = session.paragraphs('body');
      session.setParagraphAttr(authored.paraId, 'widowControl', false);
      session.setParagraphAttr(cleared.paraId, 'widowControl', null);
      saved = yrsToDocument(session, parsed);
    } finally {
      session.destroy();
    }

    const reopened = await parseDocx(await repackDocx(saved), {
      preloadFonts: false,
    });
    const paragraphs = bodyParagraphs(reopened);
    expect(paragraphs[0].formatting?.widowControl).toBe(false);
    expect(paragraphs[1].formatting?.widowControl).toBeUndefined();
  });
});
