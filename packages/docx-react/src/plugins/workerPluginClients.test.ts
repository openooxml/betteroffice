import { afterAll, expect, mock, test } from 'bun:test';
import { createPluginClients } from './createPluginClients';
import { UNAVAILABLE_DOCX_COMMANDS } from '../commands/createDocxCommandStore';
import type { PluginInvocation } from '../../../../shared/plugin-host/runtime';
import type { DocxPluginSnapshot, DocxPluginGrant } from './types';
import { createPeerlessShellController } from '../components/DocxEditor/peerless/shellController';
import { createPeerlessViewStore } from '../components/DocxEditor/peerless/viewStore';
import { createWorkerPluginGeometry, workerReadRefusal } from '../components/DocxEditor/peerless/workerPluginGeometry';
import type { PeerlessEditSessionBinding } from '../components/DocxEditor/hooks/usePeerlessEditSession';

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

test('worker reads raced by an edit or a closed document refuse instead of rejecting', async () => {
  const viewStore = createPeerlessViewStore({
    identity: { session: 'worker', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null,
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: true, canRedo: false, hyperlink: null, tableContext: null }, geometry: null,
  });
  const binding = { viewStore, admission: {}, session: { flush: async () => {} } } as unknown as PeerlessEditSessionBinding;
  const read = mock(async () => {
    const state = viewStore.getSnapshot();
    viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection: null, version: state.version + 1 });
    return { status: 'ok', version: '1', value: { ok: true, view: 'accepted', version: '1', paragraphs: [] } };
  });
  const controller = createPeerlessShellController({ read: read as never, focus() {}, zoom: 1 });
  const invocation = { refusal: () => null } as unknown as PluginInvocation<DocxPluginSnapshot>;
  const clients = createPluginClients(invocation, {
    pagedEditorRef: { current: null }, writeMode: () => 'editing', commands: () => null,
    layout: () => ({ queries: null, complete: false, failed: false }), subscribeLayout: () => () => {},
    workerRead: { version: async () => ({ ok: true, version: '0' }), readParagraphs: controller.ref.readParagraphs,
      listStories: async () => ({} as never), readStories: async () => ({} as never), findText: async () => ({} as never) },
    workerRefusal: (error) => workerReadRefusal(controller, error),
  }, () => ({ document: 'read' }), UNAVAILABLE_DOCX_COMMANDS);
  expect(await clients.read.readParagraphs({ view: 'accepted' })).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
  expect(read).not.toHaveBeenCalled();
  controller.attach(binding);
  expect(controller.getSnapshot().status).toBe('ready');
  expect(await clients.read.readParagraphs({ view: 'accepted' })).toMatchObject({ ok: false, failure: { code: 'input-failed' } });
  expect(read).toHaveBeenCalledTimes(1);
  read.mockImplementation(async () => ({ status: 'ok', version: '1', value: { ok: true, view: 'accepted', version: '1', paragraphs: [] } }));
  expect(await clients.read.readParagraphs({ view: 'accepted' })).toMatchObject({ ok: true, version: '1' });
});

test('worker refusal classification preserves unexpected errors while not ready', () => {
  const controller = createPeerlessShellController({ read: async () => ({} as never), focus() {}, zoom: 1 });
  expect(workerReadRefusal(controller, new Error('Transport disconnected'))).toBeNull();
  expect(workerReadRefusal(controller, new Error('Peer-less editor is not ready'))).toMatchObject({ failure: { code: 'document-replaced' } });
  controller.dispose();
});

for (const transition of ['edit', 'not-ready'] as const) {
  for (const focus of [false, true]) {
    test(`peer-less navigation refuses ${transition} before reveal side effects with focus ${focus}`, async () => {
      const { GlobalRegistrator } = await import('@happy-dom/global-registrator');
      if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
      const { createElement, useRef } = await import('react');
      const { act, render, cleanup } = await import('@testing-library/react');
      const { useWorkerNavigation } = await import('../components/DocxEditor/peerless/useWorkerNavigation');
      const { createDocxPluginHost } = await import('./createDocxPluginHost');
      const { defineDocxPlugin } = await import('./defineDocxPlugin');
      const identity = { session: 'navigation', incarnation: 1, epoch: 1 };
      const viewStore = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: 0,
        workerSelection: null, anchors: null, geometry: null,
        editState: { documentVersion: '0', readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
          canUndo: false, canRedo: false, hyperlink: null, tableContext: null } });
      const target = { loc: { story: 'body', paraId: 'p', offset: 0 }, position: 0 };
      const read = mock(async (request: { kind: string }) => ({ status: 'ok', version: '0',
        value: request.kind === 'readParagraphs' ? { ok: true, version: '0', paragraphs: [] } : target }));
      const focused = mock(() => {});
      const controller = createPeerlessShellController({ zoom: 1, focus: focused, read: read as never, readAt: read as never });
      const select = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }), presented: Promise.resolve('presented') } }));
      const binding = { viewStore, admission: {}, session: { flush: async () => {} }, commands: { submit: select } } as unknown as PeerlessEditSessionBinding;
      controller.attach(binding);
      let navigation!: ReturnType<typeof useWorkerNavigation>;
      function Harness() {
        const scroller = useRef<HTMLDivElement>(null);
        const host = useRef<HTMLDivElement>(null);
        navigation = useWorkerNavigation(controller, () => controller.getSnapshot().status === 'ready' ? binding : null, scroller, host, 1, {});
        return createElement('div', { ref: scroller, 'data-testid': 'scroll' }, createElement('div', { ref: host }));
      }
      const rendered = render(createElement(Harness));
      const geometry = createWorkerPluginGeometry(controller, () => null, () => null, () => {}, async (target: { paraId: string; story: string }, focus: boolean, expectVersion?: string) => {
        if (transition === 'not-ready') controller.empty();
        else {
          const state = viewStore.getSnapshot();
          viewStore.acceptWorkerState(identity, { ...state, epoch: 1, version: 1, selection: null,
            editState: { ...state.editState!, documentVersion: '1' } });
        }
        const reveal = navigation.revealParagraph as (paraId: string, story: string, focus: boolean, expectVersion?: string) => Promise<boolean>;
        return reveal(target.paraId, target.story, focus, expectVersion);
      });
      const errors: unknown[] = [];
      let context!: import('./types').DocxPluginContext<null>;
      const host = createDocxPluginHost({ pagedEditorRef: { current: null }, writeMode: () => 'editing', commands: () => null,
        layout: () => ({ queries: null, complete: false, failed: false }), subscribeLayout: () => () => {}, geometry: () => null, translate: key => key,
        workerRead: { version: async () => ({ ok: true, version: '0' }), readParagraphs: controller.ref.readParagraphs,
          listStories: async () => ({} as never), readStories: async () => ({} as never), findText: async () => ({} as never) },
        workerNavigation: geometry.navigation, workerRefusal: error => workerReadRefusal(controller, error) });
      host.setReporter(error => errors.push(error));
      host.setPlugins([defineDocxPlugin({ id: 'navigator', createState: () => null, initialize: value => { context = value; } })]);
      host.setGrants({ navigator: { document: 'read' } });
      host.openWorker('0');
      let result: unknown;
      try {
        await act(async () => {
          await context.run(async next => { result = await next.navigation.scrollToParagraph({ story: 'body', paraId: 'p' }, { expectVersion: '0', focus }); });
        });
        expect(result).toMatchObject({ ok: false, failure: { code: transition === 'edit' ? 'input-failed' : 'document-replaced' } });
        expect(rendered.getByTestId('scroll').scrollTop).toBe(0);
        expect(select).not.toHaveBeenCalled();
        expect(focused).not.toHaveBeenCalled();
        expect(errors).toEqual([]);
        controller.attach(binding);
        let ran = false;
        await context.run(async () => { ran = true; });
        expect(ran).toBe(true);
        expect(errors).toEqual([]);
      } finally { host.close('unmounted'); geometry.dispose(); cleanup(); controller.dispose(); viewStore.dispose(); }
    });
  }
}

afterAll(async () => {
  const { GlobalRegistrator } = await import('@happy-dom/global-registrator');
  if (GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});
