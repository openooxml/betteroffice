import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

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

/** The first page of a prefix layout, as the resident engine lays it out. */
function firstPage(session: YrsSession, host: YrsDocxHost): unknown {
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
  const layout = JSON.parse(
    session.layoutDocumentWithRegionsPrefixRetainedJson(JSON.stringify(request), 1)
  ) as { layout: { pages: unknown[] } };
  return layout.layout.pages[0];
}

test('a preview open lays out the first page of the document it previews', async () => {
  const full = await createYrsSession({ clientId: 98100 });
  const preview = await createYrsSession({ clientId: 98100 });
  const fullPage = firstPage(full, full.openDocx(PAGES, true));
  const previewHost = preview.openDocxPreview(PAGES, 6);
  expect(preview.paragraphs('body').length).toBeLessThan(full.paragraphs('body').length);
  expect(firstPage(preview, previewHost)).toEqual(fullPage);
  full.destroy();
  preview.destroy();
});
