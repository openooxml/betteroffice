import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { createResidentEngineSession } from './residentEngineSession';
import { saveYrsDocx } from './saveYrsDocx';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURES = resolve(import.meta.dir, '../../../../crates/docx-edit/tests/fixtures');
const DOCUMENTS = ['page-fragments/pages.docx', 'footnote-anchor.docx', 'structured-export/principal.docx'];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

function texts(session: YrsSession): Record<string, string[]> {
  return Object.fromEntries(
    session
      .storyIds()
      .sort()
      .map((story) => [story, session.paragraphs(story).map((paragraph) => paragraph.text)])
  );
}

for (const name of DOCUMENTS) {
  test(`a replica of a document the worker opened reads and saves as one opened directly: ${name}`, async () => {
    const bytes = new Uint8Array(readFileSync(resolve(FIXTURES, name)));
    const direct = await createYrsSession({ clientId: 97001 });
    const directHost = direct.openDocx(bytes, true);

    const worker = await createResidentEngineSession();
    const host = decodeDocxHostJson(worker.openDocx(bytes), bytes);
    expect(host.referencedFonts).toEqual(directHost.referencedFonts);
    expect(host.document).toEqual(directHost.document);

    const replica = await createYrsSession({ clientId: 97002 });
    replica.openDocx(bytes, false);
    replica.loadState(worker.encodeState());
    expect(texts(replica)).toEqual(texts(direct));

    const saved = await createYrsSession({ clientId: 97003 });
    saved.openDocx((await saveYrsDocx(replica)).bytes, true);
    const savedDirect = await createYrsSession({ clientId: 97004 });
    savedDirect.openDocx((await saveYrsDocx(direct)).bytes, true);
    expect(texts(saved)).toEqual(texts(savedDirect));

    for (const session of [direct, replica, saved, savedDirect]) session.destroy();
    worker.destroy();
  });
}
