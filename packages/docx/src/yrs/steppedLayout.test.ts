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

async function opened(clientId: number) {
  const session = await createYrsSession({ clientId });
  const host = session.openDocx(PAGES, true);
  return { session, request: layoutRequest(session, host) };
}

test('a region layout in steps replies as one pass and becomes the resident layout', async () => {
  const whole = await opened(98300);
  const expected = whole.session.layoutDocumentWithRegionsRetainedJson(whole.request);

  const stepped = await opened(98300);
  let progress = stepped.session.beginRegionLayout(stepped.request);
  let steps = 0;
  while (progress.layoutJson === undefined) {
    expect(progress.measuredBlocks).toBeLessThan(progress.bodyBlocks);
    progress = stepped.session.resumeRegionLayout(2);
    steps += 1;
  }
  expect(steps).toBeGreaterThan(1);
  expect(progress.layoutJson).toBe(expected);
  expect(stepped.session.residentWorkerProbe()).toEqual(whole.session.residentWorkerProbe());
  whole.session.destroy();
  stepped.session.destroy();
});

test('an edit between steps abandons the pass', async () => {
  const { session, request } = await opened(98301);
  expect(session.beginRegionLayout(request).layoutJson).toBeUndefined();
  const paragraph = session.paragraphs('body')[0]!;
  session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'x');
  expect(() => session.resumeRegionLayout(1)).toThrow();
  session.destroy();
});
