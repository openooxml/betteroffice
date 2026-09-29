import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsDocxHost, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
  )
);
const PAGES = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/docx-edit/tests/fixtures/page-fragments/pages.docx')
  )
);

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

/** A document whose second paragraph anchors a float to the margin. */
function marginFloatDocx(): Uint8Array {
  const parts: PartsMap = new Map();
  const set = (name: string, content: string) => parts.set(name, toBytes(content));
  set(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  set(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  );
  set(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapTopAndBottom/><wp:docPr id="1" name="Float"/></wp:anchor></w:drawing></w:r></w:p></w:body></w:document>'
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** The region request the editor would lay `host`'s document out with. */
function layoutRequest(session: YrsSession, host: YrsDocxHost): string {
  const font = session.registerFont(FONT);
  const request = buildResidentRegionLayoutRequest(host.document, 24, {});
  const requirements = JSON.parse(
    session.layoutFontRequirementsJson(JSON.stringify(request))
  ) as Array<{ key: string }>;
  request.measurement = {
    fontChains: Object.fromEntries(requirements.map((requirement) => [requirement.key, [font]])),
    defaults: { fontSize: 11, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
  return JSON.stringify(request);
}

/** The first page of a prefix layout, as the resident engine lays it out. */
function firstPage(session: YrsSession, host: YrsDocxHost): unknown {
  const layout = JSON.parse(
    session.layoutDocumentWithRegionsPrefixRetainedJson(layoutRequest(session, host), 1)
  ) as { layout: { pages: unknown[] } };
  return layout.layout.pages[0];
}

test('a preview open lays out the first page of the document it previews', async () => {
  const full = await createYrsSession({ clientId: 98100 });
  const preview = await createYrsSession({ clientId: 98100 });
  const fullPage = firstPage(full, full.openDocx(PAGES, true));
  const previewHost = preview.openDocxPreview(PAGES, 6)!;
  expect(preview.paragraphs('body').length).toBeLessThan(full.paragraphs('body').length);
  expect(firstPage(preview, previewHost)).toEqual(fullPage);
  // The prefix layout is the resident one a worker bootstraps from.
  expect(preview.residentWorkerSnapshot({})?.partialDocument).toBe(true);
  full.destroy();
  preview.destroy();
});

test('a preview and a worker replica of it lay out as part of a document', async () => {
  const full = await createYrsSession({ clientId: 98101 });
  const preview = await createYrsSession({ clientId: 98102 });
  const fullHost = full.openDocx(PAGES, true);
  const previewHost = preview.openDocxPreview(PAGES, 6)!;
  const layoutOf = (session: YrsSession, host: YrsDocxHost) =>
    JSON.parse(session.layoutDocumentWithRegionsRetainedJson(layoutRequest(session, host))) as {
      layout: { partial?: boolean };
    };
  expect(layoutOf(full, fullHost).layout.partial).toBeUndefined();
  expect(layoutOf(preview, previewHost).layout.partial).toBe(true);

  expect(full.residentWorkerSnapshot({})?.partialDocument).toBeUndefined();
  const snapshot = preview.residentWorkerSnapshot({})!;
  expect(snapshot.partialDocument).toBe(true);
  // What the resident worker does with a preview's snapshot.
  const replica = await createYrsSession({ clientId: 98103 });
  replica.loadState(snapshot.state);
  replica.setPartialDocument(true);
  expect(layoutOf(replica, previewHost).layout.partial).toBe(true);
  expect(replica.residentWorkerSnapshot({})?.partialDocument).toBe(true);

  // A complete open over the preview lays out the whole document again.
  for (const story of preview.storyIds()) preview.deleteStory(story);
  const reopened = preview.openDocx(PAGES, true);
  expect(layoutOf(preview, reopened).layout.partial).toBeUndefined();
  expect(preview.residentWorkerSnapshot({})?.partialDocument).toBeUndefined();
  for (const session of [full, preview, replica]) session.destroy();
});

test('a preview refuses a document with a float placed from the margin', async () => {
  const bytes = marginFloatDocx();
  const preview = await createYrsSession({ clientId: 98104 });
  expect(preview.openDocxPreview(bytes, 1)).toBeNull();
  expect(preview.storyIds()).toEqual([]);
  preview.openDocx(bytes, true);
  expect(preview.paragraphs('body').length).toBe(2);
  preview.destroy();
});
