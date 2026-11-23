import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createYrsSession, documentToYrs, saveYrsDocx } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

function makeDocxWithoutHyperlinkStyle(): Uint8Array {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '</Types>'
    )],
    ['_rels/.rels', toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`
    )],
    ['word/_rels/document.xml.rels', toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>` +
      `</Relationships>`
    )],
    ['word/styles.xml', toBytes(
      `<w:styles xmlns:w="${W}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>` +
      `</w:styles>`
    )],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Visit Example</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
    )],
  ])));
}

test('inserting a hyperlink writes Word-standard Hyperlink style and removing it cleans it up', async () => {
  const session = await createYrsSession();
  const reopenedSession = await createYrsSession();
  try {
    session.openDocx(makeDocxWithoutHyperlinkStyle(), true);
    const paraId = session.paragraphs('body')[0]!.paraId;
    const range = { story: 'body', start: { paraId, offset: 6 }, end: { paraId, offset: 13 } };

    session.setHyperlink(range, { href: 'https://example.com' });

    const saved = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    const parts = unzipContainer(saved);
    const docXml = new TextDecoder().decode(parts['word/document.xml']);
    const stylesXml = new TextDecoder().decode(parts['word/styles.xml']);

    expect(docXml).toContain('<w:hyperlink ');
    expect(docXml).toContain('<w:rStyle w:val="Hyperlink"/>');
    expect(stylesXml).toContain('w:styleId="Hyperlink"');
    expect(stylesXml).toContain('<w:color w:val="0563C1"/>');
    expect(stylesXml).toContain('<w:u w:val="single"/>');

    const reopened = await parseDocx(saved.buffer as ArrayBuffer, { preloadFonts: false });
    documentToYrs(reopenedSession, reopened);
    const linkedSegment = reopenedSession.storySegments('body').find(s => s.kind === 'text' && s.text === 'Example');
    expect(linkedSegment?.attributes.runStyle).toEqual({ styleId: 'Hyperlink' });
    expect(linkedSegment?.attributes.textColor).toMatchObject({ rgb: '0563C1', inheritedHyperlink: true });
    expect(linkedSegment?.attributes.underline).toMatchObject({ style: 'single', inheritedHyperlink: true });

    // Explicitly toggling underline off on a link whose source lacked Hyperlink style emits <w:u w:val="none"/>.
    session.toggleMark(range, { type: 'underline' });
    const noUnderlineSaved = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    const noUnderlineXml = new TextDecoder().decode(unzipContainer(noUnderlineSaved)['word/document.xml']);
    expect(noUnderlineXml).toContain('<w:rStyle w:val="Hyperlink"/>');
    expect(noUnderlineXml).toContain('<w:u w:val="none"/>');

    // Unlinking removes <w:hyperlink> and clears the default Hyperlink style.
    session.setHyperlink(range, null);
    const unlinkedSaved = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    const unlinkedXml = new TextDecoder().decode(unzipContainer(unlinkedSaved)['word/document.xml']);
    expect(unlinkedXml).not.toContain('<w:hyperlink');
    expect(unlinkedXml).not.toContain('w:val="Hyperlink"');
  } finally {
    session.destroy();
    reopenedSession.destroy();
  }
}, 30_000);
