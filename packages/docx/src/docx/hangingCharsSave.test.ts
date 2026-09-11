import { describe, expect, it } from 'bun:test';
import { parseDocx } from '.';
import { repackDocx } from './rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from './rezip/parts';
import { readDocxContainer } from './zipContainer';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function document(ind: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:ind ${ind}/></w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;
}

async function browserSave(ind: string): Promise<string> {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/document.xml', toBytes(document(ind)));
  const bytes = rezipPartsToArrayBuffer(parts);
  const parsed = await parseDocx(bytes, { preloadFonts: false });
  return readDocxContainer(await repackDocx(parsed)).text('word/document.xml') ?? '';
}

describe('character-unit hanging indents through the browser save', () => {
  // the model crosses JSON on this path, and JSON has no negative zero
  it('keeps hangingChars="0" as hanging', async () => {
    const xml = await browserSave('w:hangingChars="0"');
    expect(xml).toContain('w:hangingChars="0"');
    expect(xml).not.toContain('w:firstLineChars');
  });

  it('keeps a hanging and a first-line character indent apart', async () => {
    expect(await browserSave('w:hangingChars="150"')).toContain('w:hangingChars="150"');
    expect(await browserSave('w:firstLineChars="200"')).toContain('w:firstLineChars="200"');
  });
});
