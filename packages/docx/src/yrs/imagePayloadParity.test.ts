import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='),
  (char) => char.charCodeAt(0)
);

const RULE =
  '<w:pict><v:rect id="rule" style="width:487.05pt;height:2.85pt" o:hrpct="0" o:hralign="center" o:hr="t">' +
  '<v:imagedata r:id="rIdImage" o:title="rule"/></v:rect></w:pict>';

const PICTURE =
  '<w:drawing><wp:inline distT="12345" distB="6789" distL="114300" distR="100000">' +
  '<wp:extent cx="1234567" cy="765433"/><wp:effectExtent l="1000" t="2000" r="3000" b="4000"/>' +
  '<wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1234567" cy="765433"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';

/** A VML rule picture and a DrawingML picture, both sized off the 96 dpi pixel grid, then a text paragraph. */
function picturesPackage(): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${OFFICE}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${OFFICE}/image" Target="media/image1.png"/></Relationships>`)],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${OFFICE}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r>${RULE}</w:r></w:p><w:p><w:r>${PICTURE}</w:r></w:p><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)],
    ['word/media/image1.png', PNG],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function imagePayloads(session: YrsSession): Record<string, unknown>[] {
  return session.storySegments('body').flatMap((segment) =>
    segment.kind === 'embed' && segment.embedKind === 'image' ? [segment.payload] : []
  );
}

describe('picture payload parity', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('projects the same picture payloads as the native seed', async () => {
    const bytes = picturesPackage();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const engine = await createYrsSession({ clientId: 47101 });
    const projected = await createYrsSession({ clientId: 47102 });
    try {
      engine.seedFromDocx(bytes);
      documentToYrs(projected, parsed);
      const native = imagePayloads(engine);
      expect(native).toHaveLength(2);
      expect([native[0].width, native[0].height].map((px) => Number.isInteger(px))).toEqual([false, false]);
      expect(native[1]).toMatchObject({ width: 1234567 / 914400 * 96, distTop: 12345 / 914400 * 96, effectExtentLeft: 1000 / 914400 * 96 });
      expect(imagePayloads(projected)).toEqual(native);
    } finally {
      engine.destroy();
      projected.destroy();
    }
  });

  it('keeps untouched pictures byte-identical on a hydrated save', async () => {
    const bytes = picturesPackage();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const projected = await createYrsSession({ clientId: 47103 });
    const hydrated = await createYrsSession({ clientId: 47104 });
    const debug = console.debug;
    const diagnostics: string[] = [];
    console.debug = (...args: unknown[]) => {
      diagnostics.push(args.map((value) => (typeof value === 'string' ? value : JSON.stringify(value))).join(' '));
    };
    try {
      documentToYrs(projected, parsed);
      hydrated.openDocx(bytes, false);
      hydrated.loadState(projected.encodeState());
      const text = hydrated.paragraphs('body')[2];
      hydrated.insertText({ story: 'body', paraId: text.paraId, offset: 1 }, 'y');
      const body = hydrated.paragraphSavePlan().splicedParts?.find((entry) => entry.part.replace(/^\//, '') === 'word/document.xml');
      expect(body?.changed).toHaveLength(1);
      const xml = new TextDecoder().decode(unzipContainer((await saveYrsDocx(hydrated)).bytes)['word/document.xml']);
      expect(xml).toContain(RULE);
      expect(xml).toContain(PICTURE);
      expect(xml).toContain('xy');
      expect(diagnostics.filter((message) => message.includes('"drawing"'))).toEqual([]);
    } finally {
      console.debug = debug;
      projected.destroy();
      hydrated.destroy();
    }
  });
});
