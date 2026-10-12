import { grantsEditBatch } from '../../../../shared/plugin-host/grants';
import type { PluginInvocation } from '../../../../shared/plugin-host/runtime';
import { pluginRefusal, unlessAborted, type PptxPluginClients } from '../plugins/createPluginClients';
import type { PptxPluginGrant, PptxPluginSnapshot, PptxPluginRefusal, PptxPluginNavigationResult, PptxPluginNavigationOptions } from '../plugins/types';
import type { PptxCommandStore } from '../commands/types';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import type { PptxAsyncEditorApi } from './createPeerlessEditorApi';
import type { PresentationCommandAck, PresentationCommandIntent } from '@betteroffice/pptx';
import { peerlessGestureActive } from './peerlessViewStore';

export function peerlessPluginClients(invocation: PluginInvocation<PptxPluginSnapshot>, run: PeerlessPresentation,
  api: () => PptxAsyncEditorApi | null, grant: () => PptxPluginGrant, store: PptxCommandStore,
  commit: (intent: PresentationCommandIntent) => Promise<PresentationCommandAck>): PptxPluginClients {
  const denial = () => {
    const reason = invocation.refusal();
    return reason ? pluginRefusal(reason) : !run.current ? pluginRefusal('document-replaced')
      : peerlessGestureActive(run.store) ? pluginRefusal('input-failed', 'An editor gesture is active') : null;
  };
  async function admitted<T>(operation: () => Promise<T>, writing = false): Promise<T | PptxPluginRefusal> {
    const before = denial();
    if (before) return before;
    try {
      let started = false;
      const result = await unlessAborted([invocation.signal, invocation.lifetimeSignal], (async () => {
        await run.store.fence();
        const refused = denial();
        if (refused) return refused;
        run.require();
        started = true;
        const value = await operation();
        return writing ? value : denial() ?? value;
      })(), () => started);
      return result ?? denial() ?? pluginRefusal('aborted');
    } catch {
      return denial() ?? pluginRefusal('input-failed');
    }
  }
  const editor = () => { const value = api(); if (!value) throw new Error('Editor unavailable'); return value; };
  const read: PptxPluginClients['read'] = {
    version: () => admitted(async () => ({ ok: true as const, version: (await run.client.flush()).version })),
    readContent: (request) => admitted(async () => (await run.client.read(request)).result),
    findText: (request) => admitted(() => run.client.access!('findText', request)),
    validateEdits: (request) => admitted(async () => run.store.getSnapshot().editState.readOnly
      ? pluginRefusal('read-only') : run.client.access!('validateEdits', request)),
  };
  const edits: PptxPluginClients['edits'] = grantsEditBatch(grant(), 'separate') ? {
    applyEdits: (request) => admitted(async () => {
      const refused = denial();
      if (refused) return refused;
      if (!grantsEditBatch(grant(), request.history)) return pluginRefusal('permission-denied');
      if (run.store.getSnapshot().editState.readOnly) return pluginRefusal('read-only');
      const ack = await commit({ kind: 'applyEdits',
        payload: { request, inputTime: performance.now(), undoGroup: 'plugin' } });
      return ack.editResult ?? pluginRefusal('input-failed');
    }, true),
  } : null;
  const commands = {
    getDescriptor: (id: never) => store.getDescriptor(id), getState: (id: never, args?: never) => store.getState(id, args),
    subscribe: store.subscribe, execute: (id: never, args: never) => admitted(() => store.execute(id, args), true),
  } as unknown as PptxPluginClients['commands'];
  async function navigate(options: PptxPluginNavigationOptions, operation: () => Promise<boolean>): Promise<PptxPluginNavigationResult> {
    return admitted(async () => {
      if (run.client.state.version !== options.expectVersion) return { ok: false as const, failure: { code: 'stale-version' as const, message: 'The presentation changed' } };
      const result = await operation();
      if (options.focus && result) editor().focus();
      return result ? { ok: true as const } : { ok: false as const, failure: { code: 'missing-target' as const, message: 'The target is unavailable' } };
    });
  }
  const navigation: PptxPluginClients['navigation'] = {
    goToSlide: (target, options) => navigate(options, () => editor().goToSlide(run.client.state.slides.findIndex((slide) => slide.id === target.slideId) + 1)),
    selectText: (target, options) => navigate(options, () => editor().selectText({ ...target,
      slide: run.client.state.slides.findIndex((slide) => slide.id === target.slideId) + 1, start: target.anchor, end: target.focus })),
    selectShape: (target, options) => navigate(options, async () => {
      const index = run.client.state.slides.findIndex((slide) => slide.id === target.slideId);
      if (index < 0 || !await editor().goToSlide(index + 1)) return false;
      if (denial()) return false;
      const frame = run.store.getSnapshot().confirmed;
      if (!frame) return false;
      const ack = await run.store.dispatch({ kind: 'selectObject', payload: { frame: frame.frame,
        shapeId: target.shapeId,
        inputTime: performance.now(), undoGroup: 'plugin-selection' } });
      return !ack.outcomes.some((outcome) => outcome.status === 'refused') && run.store.getSnapshot().selection?.shapeId === target.shapeId;
    }),
  };
  return { read, edits, commands, navigation };
}
