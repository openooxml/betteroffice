import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes } from '../../docx/rezip/parts';
import { preloadEditWasm } from '../../wasm/edit';
import { createYrsSession } from '../../yrs';
import type { DisplayList } from './displayList';
import { buildMirrorPage } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm')
  )));
}, 30_000);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function docx(eastAsia: string): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  add('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="office" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`);
  add('word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
    `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>` +
    `<w:lang w:val="en-US" w:eastAsia="${eastAsia}" w:bidi="ar-SA"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`);
  add('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` +
    '<w:p><w:r><w:t>Annual report</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>年次報告</w:t></w:r></w:p>' +
    '<w:sectPr/></w:body></w:document>');
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function paragraphLanguages(eastAsia: string): Promise<[string, string][]> {
  const session = await createYrsSession({ clientId: 97900 });
  try {
    session.openDocx(docx(eastAsia), true);
    const layout = session.layoutDocumentWithRegionsJson(JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: {
        fontChains: {},
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: false,
      },
      renderEnv: {},
    }));
    const list = JSON.parse(session.buildDisplayListJson(layout)) as DisplayList;
    const mirror = buildMirrorPage(list.pages[0]!);
    return Array.from(mirror.querySelectorAll<HTMLElement>('.layout-paragraph'))
      .map((paragraph) => [paragraph.textContent ?? '', paragraph.lang] as [string, string])
      .filter(([text]) => text.trim() !== '');
  } finally {
    session.destroy();
  }
}

test('English text in a document with an East Asian default keeps its Latin language for screen readers', async () => {
  for (const eastAsia of ['zh-CN', 'ja-JP']) {
    expect(await paragraphLanguages(eastAsia)).toEqual([
      ['Annual report', 'en-US'],
      ['年次報告', eastAsia],
    ]);
  }
});
