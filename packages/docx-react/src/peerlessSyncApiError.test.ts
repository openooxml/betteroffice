import { expect, test } from 'bun:test';
import { DOCX_REF_ASYNC_TWINS, DOCX_REF_ASYNC_TWIN_EXEMPTIONS } from './components/DocxEditor/hooks/useDocxEditorRefApi';
import { createPeerlessShellController } from './components/DocxEditor/peerless/shellController';
import type { DocxEditorRef } from './index';
import { PeerlessSyncApiError } from './index';

test('synchronous ref members of the default editor throw the exported PeerlessSyncApiError', () => {
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
    expect(thrown).toBeInstanceOf(PeerlessSyncApiError);
    expect((thrown as Error).name).toBe('PeerlessSyncApiError');
    expect((thrown as Error).message).toContain(member);
  }
});
