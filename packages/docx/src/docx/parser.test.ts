import { expect, test } from 'bun:test';
import { parseDocx } from './parser';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';

test('foreign packages expose the NotAWordDocument parse error code', async () => {
  for (const parts of [
    new Map([['readme.txt', toBytes('Archive contents')]]),
    new Map([['word/document.xml', toBytes('<document xmlns="urn:foreign"><body/></document>')]]),
  ]) {
    await expect(parseDocx(rezipPartsToArrayBuffer(parts), { preloadFonts: false }))
      .rejects.toMatchObject({
        code: 'NotAWordDocument',
        message: expect.stringMatching(/^Failed to parse DOCX: not a Word document: /),
      });
  }
});

test('malformed Word XML keeps the MalformedXml parse error code', async () => {
  const parts = new Map([
    ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body></w:document>')],
  ]);
  await expect(parseDocx(rezipPartsToArrayBuffer(parts), { preloadFonts: false }))
    .rejects.toMatchObject({ code: 'MalformedXml', part: 'word/document.xml', offset: expect.any(Number) });
});
