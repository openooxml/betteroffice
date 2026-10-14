import { expect, mock, test } from 'bun:test';
import { createPluginClients } from './createPluginClients';
import { UNAVAILABLE_DOCX_COMMANDS } from '../commands/createDocxCommandStore';
import type { PluginInvocation } from '../../../../shared/plugin-host/runtime';
import type { DocxPluginSnapshot, DocxPluginGrant } from './types';

test('worker clients validate and recheck revoked grants after pending input', async () => {
  let grant: DocxPluginGrant = { document: 'write', editBatches: true };
  let refused: 'aborted' | null = null;
  let authorize!: () => unknown;
  let finish!: () => void;
  const invocation = { refusal: () => refused } as unknown as PluginInvocation<DocxPluginSnapshot>;
  const validation = mock(async () => ({ ok: true as const, version: 'v1', baseVersion: 'v1', wouldApply: false, previews: [] }));
  const clients = createPluginClients(invocation, {
    pagedEditorRef: { current: null }, writeMode: () => 'editing', commands: () => null,
    layout: () => ({ queries: null, complete: false, failed: false }), subscribeLayout: () => () => {},
    workerRead: { version: async () => ({ ok: true, version: 'v1' }), readParagraphs: async () => ({} as never),
      listStories: async () => ({} as never), readStories: async () => ({} as never), findText: async () => ({} as never), validateEdits: validation },
    workerEdits: { applyEdits: async (_request, check) => { authorize = check; await new Promise<void>(resolve => { finish = resolve; }); return check() ?? ({} as never); } },
    workerNavigation: { scrollToParagraph: async () => ({ ok: true }) },
  }, () => grant, UNAVAILABLE_DOCX_COMMANDS);
  expect(await clients.read.validateEdits({ expectVersion: 'v1', steps: [] })).toMatchObject({ ok: true });
  const pending = clients.edits!.applyEdits({ expectVersion: 'v1', steps: [] });
  grant = { document: 'write' };
  expect(authorize()).toMatchObject({ failure: { code: 'permission-denied' } });
  finish();
  expect(await pending).toMatchObject({ failure: { code: 'permission-denied' } });
  expect(await clients.navigation.scrollToParagraph({ story: 'body', paraId: 'p' }, { expectVersion: 'v1' })).toEqual({ ok: true });
  refused = 'aborted';
  expect(await clients.read.validateEdits({ expectVersion: 'v1', steps: [] })).toMatchObject({ failure: { code: 'aborted' } });
  expect(validation).toHaveBeenCalledTimes(1);
});

test('worker plugin edits return a document replacement refusal when a pending batch rejects', async () => {
  let replaced = false;
  let reject!: (error: Error) => void;
  const invocation = { refusal: () => replaced ? 'document-replaced' : null } as unknown as PluginInvocation<DocxPluginSnapshot>;
  const clients = createPluginClients(invocation, {
    pagedEditorRef: { current: null }, writeMode: () => 'editing', commands: () => null,
    layout: () => ({ queries: null, complete: false, failed: false }), subscribeLayout: () => () => {},
    workerRead: { version: async () => ({ ok: true, version: 'v1' }), readParagraphs: async () => ({} as never),
      listStories: async () => ({} as never), readStories: async () => ({} as never), findText: async () => ({} as never) },
    workerEdits: { applyEdits: () => new Promise((_resolve, fail) => { reject = fail; }) },
  }, () => ({ document: 'write', editBatches: true }), UNAVAILABLE_DOCX_COMMANDS);
  const pending = clients.edits!.applyEdits({ expectVersion: 'v1', steps: [] });
  replaced = true;
  reject(new Error('Document was replaced'));
  expect(await pending).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
});
