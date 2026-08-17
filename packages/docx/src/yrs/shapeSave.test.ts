import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { Document } from '../types/document';
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
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const TEXT_BOX = `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="41" name="Text Box 41"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>BARE-BODY</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    ${TEXT_BOX}
    <w:p><w:r><w:t>BODY-TEXT</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

function fixture(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/document.xml', toBytes(DOCUMENT));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

interface SaveOptions {
  seeder?: 'native' | 'projected';
  edit?: (session: YrsSession) => void;
}

/** Seed `bytes` the editor's way, run the edit, project and repack. */
async function saveBody(bytes: Uint8Array, options: SaveOptions = {}): Promise<string> {
  const { seeder = 'native', edit } = options;
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const session = await createYrsSession({ clientId: 53101 });
  let saved: Document;
  try {
    if (seeder === 'native') session.seedFromDocx(bytes);
    else documentToYrs(session, parsed);
    edit?.(session);
    saved = yrsToDocument(session, parsed);
  } finally {
    session.destroy();
  }
  return readDocxContainer(await repackDocx(saved)).text('word/document.xml') ?? '';
}

function drawing(xml: string): string {
  const match = /<w:drawing>[\s\S]*?<\/w:drawing>/.exec(xml);
  return match?.[0] ?? '';
}

function appendToLastParagraph(session: YrsSession, text: string): void {
  const spans = session.paragraphSpans('body');
  const last = spans[spans.length - 1];
  session.insertText({ story: 'body', paraId: last.paraId, offset: last.length }, text);
}

describe('shape save projection', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it.each(['native', 'projected'] as const)(
    'keeps a text box body through an untouched %s projection',
    async (seeder) => {
      const xml = await saveBody(fixture(), { seeder });

      expect(xml.match(/<w:txbxContent>/g)?.length).toBe(1);
      expect(xml).toContain('BARE-BODY');
      expect(drawing(xml)).toContain('<wps:cNvSpPr txBox="1"/>');
    }
  );

  it('keeps a text box body when the body text around it is edited', async () => {
    const untouched = await saveBody(fixture());
    const edited = await saveBody(fixture(), {
      edit: (session) => appendToLastParagraph(session, ' EDITED-IN-YRS'),
    });

    expect(edited).toContain('BODY-TEXT EDITED-IN-YRS');
    expect(edited).toContain('BARE-BODY');
    expect(drawing(edited)).toBe(drawing(untouched));
  });
});
