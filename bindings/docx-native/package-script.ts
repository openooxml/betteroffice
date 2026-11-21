import JSZip from 'jszip';
import type * as Core from '../../packages/docx/src/core';
import type * as Yrs from '../../packages/docx/src/yrs';

export async function documentFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2024-01-01T00:00:00Z');
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>', { date });
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>', { date });
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Package text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>', { date });
  zip.file('custom/untouched.bin', new Uint8Array([0, 127, 255]), { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

export async function exercisePackage(yrs: typeof Yrs, core: typeof Core, bytes: Uint8Array) {
  const session = await yrs.createYrsSession({ clientId: 7 });
  const reopened = await yrs.createYrsSession({ clientId: 8 });
  try {
    session.seedFromDocx(bytes);
    const first = session.paragraphs('body')[0]!;
    const before = first.text;
    session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Native ');
    session.insertText({ story: 'body', paraId: first.paraId, offset: 7 }, 'Ω ');
    const texts = session.paragraphs('body').map(({ text }) => text);
    const saved = await yrs.saveYrsDocx(session, { updateModifiedDate: false });
    reopened.seedFromDocx(saved.bytes);
    const reopenedTexts = reopened.paragraphs('body').map(({ text }) => text);
    const parsed = await core.parseDocx(saved.bytes);
    return {
      bytes: saved.bytes,
      results: {
        before,
        texts,
        reopenedTexts,
        blocks: parsed.package.document.content.length,
        markdown: await core.exportDocxMarkdown(saved.bytes, { revisionView: 'accepted' }),
        preflight: await core.fontPreflight(saved.bytes),
      },
    };
  } finally {
    reopened.destroy();
    session.destroy();
  }
}
