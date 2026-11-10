import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { openDocx } from '../src';
import { renderDocxPage } from '../src/render';
import fixture from '../../../crates/docx-edit/tests/fixtures/behind_objects.json';

/** A page with one anchored picture whose bytes are `metafile`. */
async function pictureDocx(metafile: Uint8Array): Promise<Uint8Array> {
  const image = fixture.image.replace('IMAGE_RANK', '0').replace('IMAGE_BEHIND', '0');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/x-emf"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${fixture.namespaces}><w:body>${image}${fixture.body}</w:body></w:document>`);
  zip.file('word/media/image.png', metafile);
  return zip.generateAsync({ type: 'uint8array' });
}

test('previews draw metafile pictures from the document copy as the live session does', async () => {
  for (const name of ['clip-bitmap.emf', 'text.emf']) {
    const picture = new Uint8Array(readFileSync(new URL(`../../../crates/ooxml-metafile/tests/fixtures/${name}`, import.meta.url)));
    const doc = await openDocx(await pictureDocx(picture), { renderer: renderDocxPage });
    try {
      const live = await renderDocxPage(doc.session, 1);
      const preview = await doc.render(1);
      expect(Buffer.from(preview.png).equals(Buffer.from(live.png))).toBe(true);
    } finally { doc.close(); }
  }
});
