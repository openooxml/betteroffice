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
import { captureSessionSave, saveYrsDocx, writeSessionSave } from './saveYrsDocx';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='),
  (char) => char.charCodeAt(0)
);

function picture(inline: string, effectExtent: string): string {
  return (
    `<w:p><w:r><w:drawing>${inline}<wp:extent cx="2194560" cy="822960"/>${effectExtent}` +
    '<wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2194560" cy="822960"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
  );
}

function fixture(inline: string, effectExtent: string): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${OFFICE}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/_rels/document.xml.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${OFFICE}/image" Target="media/image1.png"/></Relationships>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>' +
        '<w:p><w:r><w:t>Before</w:t></w:r></w:p>' +
        picture(inline, effectExtent) +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1152" w:right="1224" w:bottom="1152" w:left="1224"/></w:sectPr></w:body></w:document>'
    )
  );
  parts.set('word/media/image1.png', PNG);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

type Seeder = 'native' | 'projected';
type Save = 'editor' | 'repack' | 'edited';

async function saveOnce(bytes: Uint8Array, seeder: Seeder, save: Save): Promise<Uint8Array> {
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const session: YrsSession = await createYrsSession({ clientId: 66101 });
  try {
    let base: Document;
    if (seeder === 'native') {
      session.seedFromDocx(bytes);
      base = session.materializeDocx()!;
    } else {
      documentToYrs(session, parsed);
      base = parsed;
    }
    if (save === 'repack') return new Uint8Array(await repackDocx(yrsToDocument(session, base)));
    if (save === 'edited') {
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      if (seeder === 'native') return (await saveYrsDocx(session)).bytes;
      return new Uint8Array(await repackDocx(yrsToDocument(session, base)));
    }
    const capture = captureSessionSave(session);
    const saved = await writeSessionSave(
      session,
      yrsToDocument(session, base),
      capture,
      base.originalBuffer ?? (bytes.buffer as ArrayBuffer),
      {},
      () => false
    );
    return saved.bytes;
  } finally {
    session.destroy();
  }
}

function drawingTags(bytes: Uint8Array): { inline: string; effectExtent: string | null } {
  const xml = readDocxContainer(bytes.buffer as ArrayBuffer).text('word/document.xml') ?? '';
  return {
    inline: /<wp:inline\b[^>]*>/.exec(xml)?.[0] ?? '',
    effectExtent: /<wp:effectExtent\b[^>]*\/>/.exec(xml)?.[0] ?? null,
  };
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

const CASES = [
  { name: 'absent attributes', inline: '<wp:inline>', effectExtent: '' },
  {
    name: 'Word zero attributes',
    inline: '<wp:inline distT="0" distB="0" distL="0" distR="0">',
    effectExtent: '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
  },
  {
    name: 'partial attributes',
    inline: '<wp:inline distL="114300" distR="114300">',
    effectExtent: '<wp:effectExtent l="19050" t="0" r="0" b="0"/>',
  },
] as const;

describe('inline picture wrap distances and effect extent', () => {
  for (const seeder of ['native', 'projected'] as const) {
    for (const save of ['editor', 'repack', 'edited'] as const) {
      for (const { name, inline, effectExtent } of CASES) {
        it(`${seeder} ${save} save keeps ${name} through three cycles`, async () => {
          let bytes = fixture(inline, effectExtent);
          for (let cycle = 0; cycle < 3; cycle += 1) {
            bytes = await saveOnce(bytes, seeder, save);
            expect(drawingTags(bytes)).toEqual({
              inline,
              effectExtent: effectExtent || null,
            });
          }
        });
      }
    }
  }
});
