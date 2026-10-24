import { expect, test } from 'bun:test';
import { DOCX_REF_ASYNC_TWINS, DOCX_REF_ASYNC_TWIN_EXEMPTIONS } from './components/DocxEditor/hooks/useDocxEditorRefApi';
import { createPeerlessShellController } from './components/DocxEditor/peerless/shellController';
import type { DocxEditorRef } from './index';
import { DocxSyncApiError } from './index';

test('synchronous ref members of the default editor throw the exported DocxSyncApiError', () => {
  const { ref } = createPeerlessShellController({ read: async () => ({} as never), focus() {}, zoom: 1 });
  const members = [...Object.keys(DOCX_REF_ASYNC_TWINS), ...DOCX_REF_ASYNC_TWIN_EXEMPTIONS] as (keyof DocxEditorRef)[];
  expect(members.length).toBeGreaterThan(0);
  for (const member of members) {
    let thrown: unknown;
    try {
      Reflect.apply(ref[member] as (...args: unknown[]) => unknown, ref, []);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(DocxSyncApiError);
    expect((thrown as Error).name).toBe('DocxSyncApiError');
    expect(thrown).toMatchObject({ code: 'sync-api-unavailable', member });
    expect((thrown as Error).message).toContain(member);
    expect((thrown as Error).message).not.toMatch(/worker-only|peer-?less/i);
  }
});
