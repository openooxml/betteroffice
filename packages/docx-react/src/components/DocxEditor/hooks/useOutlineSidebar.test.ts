import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createYrsSession } from '@betteroffice/docx/yrs';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { PagedEditorRef } from '../PagedEditor';
import { collectYrsHeadings } from './useOutlineSidebar';

const ROOT = resolve(import.meta.dir, '../../../../../..');

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);

test('the outline lists the headings the engine classifies', async () => {
  const session = await createYrsSession({ clientId: 91001 });
  try {
    session.openDocx(
      new Uint8Array(
        readFileSync(
          resolve(ROOT, 'crates/docx-edit/tests/fixtures/structured-export/principal.docx')
        )
      ),
      true
    );
    const editor = {
      getYrsSession: () => session,
      yrsLocToDisplayPosition: () => 3,
    } as unknown as PagedEditorRef;
    expect(collectYrsHeadings(editor).map(({ text, level }) => [text, level])).toEqual([
      ['Structured export', 0],
      ['Direct outline', 1],
      ['Inherited from Heading1', 0],
      ['Built-in style fallback', 1],
      ['Numbered heading', 0],
    ]);
  } finally {
    session.destroy();
  }
});
