import { expect, mock, test } from 'bun:test';
import type { ComponentProps } from 'react';
import { PptxPeerNotReadyError, PptxWorkerEditorDisposedError, PptxWorkerEditorFailedError } from '@betteroffice/pptx';
import type { PptxEditRequest, PptxWorkerEditorAccess, PptxWorkerEditorSession, PptxWorkerEditorState } from '@betteroffice/pptx';
import { PptxEditor } from '../PptxEditor';
import type { PptxEditorApi, PptxEditorProps, PptxWorkerViewerApi } from '../PptxEditor';
import { createWorkerEditorApi, type PptxWorkerEditorApi } from './createWorkerEditorApi';
import type { PptxWorkerEditorProps } from './PptxWorkerEditor';

type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2) ? true : false;
function assertType<T extends true>(_value?: T) {}

function setup(gestureActive: () => boolean = () => false) {
  let state: PptxWorkerEditorState = { stage: 'ready', sequence: 0, acknowledgedSequence: 0 };
  const peer = {
    version: mock(() => 'v1'), readContent: mock(() => ({ ok: true, version: 'v1', slides: [], stories: [] })),
    findText: mock(() => ({ ok: true, version: 'v1', matches: [], truncated: false })),
    validateEdits: mock(() => ({ ok: true, version: 'v1' })),
    applyEdits: mock(() => ({ ok: true, version: 'v1', applied: 1 })),
    insertText: mock(() => {}),
  } as unknown as PptxWorkerEditorAccess;
  const session = {
    get state() { return state; },
    get hydrated() { return state.stage === 'ready'; },
    get failure() { return state.failure; },
    whenHydrated: mock(async () => {}), handleAsync: mock(async () => peer),
    saveAsync: mock(async () => new Uint8Array([2])),
    recoverySave: mock(async () => ({ bytes: new Uint8Array([3]), recovery: true as const })),
  } as unknown as PptxWorkerEditorSession;
  const ui = {
    commands: { execute: mock(async () => ({ ok: true, status: 'executed' })) },
    applyEdits: mock(async (request: PptxEditRequest) => peer.applyEdits(request)),
    flushPendingInput: mock(async () => {}), refresh: mock(() => {}), refreshProposals: mock(() => {}),
    clearSelection: mock(() => {}), focus: mock(() => {}), goToSlide: mock(() => true),
    getPositionAtPoint: mock(() => null), selectText: mock(() => true),
  } as unknown as Omit<PptxEditorApi, 'handle' | 'save'>;
  const api = createWorkerEditorApi(session, ui, () => peer, async () => true, gestureActive);
  return { api, peer, session, ui, stage: (stage: PptxWorkerEditorState['stage']) => { state = { ...state, stage }; },
    fail: () => { state = { ...state, stage: 'failed', failure: new PptxWorkerEditorFailedError(new Error('lost')) }; } };
}

test('legacy_contextual_api_types_are_unchanged', () => {
  assertType<Equal<Parameters<typeof PptxEditor>[0], PptxEditorProps>>();
  assertType<Equal<ComponentProps<typeof PptxEditor>, PptxEditorProps>>();
  const editor: PptxEditorProps = { fonts: [], onReady(api) { assertType<Equal<typeof api, PptxEditorApi>>(); } };
  const viewer = PptxEditor({ fonts: [], readOnly: true, experimentalWorkerOpen: true, onReady(api) {
    assertType<Equal<typeof api, PptxWorkerViewerApi>>();
  } });
  const worker: PptxWorkerEditorProps = { fonts: [], experimentalWorkerOpen: true,
    onReady(api) { assertType<Equal<typeof api, PptxWorkerEditorApi>>(); } };
  expect(editor.fonts).toEqual(worker.fonts);
  expect(viewer).toBeDefined();
});

test('worker_sync_save_and_handle_are_null', async () => {
  const { api, session } = setup();
  expect(api.handle).toBeNull();
  expect(api.save()).toBeNull();
  expect(await api.saveAsync()).toEqual(new Uint8Array([2]));
  expect(session.saveAsync).toHaveBeenCalledTimes(1);
  expect('flush' in api).toBe(false);
  expect('selectTextAsync' in api).toBe(false);
  const gesture = setup(() => true);
  await expect(gesture.api.saveAsync()).rejects.toThrow('Finish the pointer gesture');
  await expect(gesture.api.applyEdits({ expectVersion: 'v1', steps: [] })).rejects.toThrow('Finish the pointer gesture');
  expect(gesture.session.saveAsync).not.toHaveBeenCalled();
  expect(gesture.peer.applyEdits).not.toHaveBeenCalled();
});

test('async_access_never_exposes_raw_peer', async () => {
  const { api, peer, ui } = setup();
  expect(await api.handleAsync()).toBe(peer);
  expect('save' in await api.handleAsync()).toBe(false);
  expect('dispose' in await api.handleAsync()).toBe(false);
  expect(await api.readContent()).toEqual({ ok: true, version: 'v1', slides: [], stories: [] });
  expect(ui.flushPendingInput).toHaveBeenCalled();
  expect(peer.readContent).toHaveBeenCalledTimes(1);
});

test('applyEdits_before_hydration_is_not_queued', async () => {
  const { api, peer, ui, stage } = setup();
  stage('hydrating');
  await expect(api.applyEdits({ expectVersion: 'v1', steps: [] })).rejects.toBeInstanceOf(PptxPeerNotReadyError);
  stage('ready');
  expect(peer.applyEdits).not.toHaveBeenCalled();
  const pending = api.applyEdits({ expectVersion: 'v1', steps: [] });
  expect(ui.applyEdits).toHaveBeenCalledTimes(1);
  expect(peer.applyEdits).toHaveBeenCalledTimes(1);
  await pending;
});

test('retired_api_only_retains_recovery_access', async () => {
  const { api, stage, fail } = setup();
  const commands = api.commands;
  fail();
  expect(api.failure).toBeInstanceOf(PptxWorkerEditorFailedError);
  expect(api.hydrated).toBe(false);
  expect((await api.recoverySave()).bytes).toEqual(new Uint8Array([3]));
  stage('disposed');
  expect(() => commands.execute('bold', null)).toThrow(PptxWorkerEditorDisposedError);
  for (const member of ['handle', 'commands', 'hydrated', 'failure'] as const)
    expect(() => api[member]).toThrow(PptxWorkerEditorDisposedError);
  for (const operation of [api.save, api.refresh, api.refreshProposals, api.clearSelection, api.focus,
    () => api.goToSlide(1), () => api.selectText({ slide: 1, shapeId: 'shape', storyId: 'story', start: 0, end: 0 }),
    () => api.getPositionAtPoint(0, 0)]) expect(operation).toThrow(PptxWorkerEditorDisposedError);
  for (const operation of [api.saveAsync, api.whenHydrated, api.handleAsync, api.flushPendingInput,
    api.version, api.readContent, () => api.goToSlideAsync(1), () => api.applyEdits({ expectVersion: 'v1', steps: [] })])
    await expect(operation()).rejects.toBeInstanceOf(PptxWorkerEditorDisposedError);
  expect(await api.recoverySave()).toEqual({ bytes: new Uint8Array([3]), recovery: true });
});
