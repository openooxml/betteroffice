import { expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { act, render, waitFor } from '@testing-library/react';
import { useCallback, useEffect, useState } from 'react';
import type { AsyncCollaborationReplica } from '@betteroffice/pptx';
import { CollaborationProvider, type CollaborationTransportEvent } from '@betteroffice/pptx/editor';
import { open, story } from './worker/peerlessTestHarness';
import * as wasm from '../../pptx/src/wasm/loader';
import { workerEditorSessionOpener } from './worker/useEditableSessionPresentation';
import { peerlessSessionOpener } from './worker/usePeerlessPresentation';
import { PptxEditor } from './PptxEditor';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';
import { resetPptxEngineWarningsForTests } from './engineWarnings';

test('default readiness uses one worker engine and no eager main-thread engine', async () => {
  const hybrid = spyOn(workerEditorSessionOpener, 'open');
  const engines = spyOn(wasm, 'openPresentation');
  const warn = spyOn(console, 'warn');
  try {
    const { api, view } = await open();
    expect(view.getByTestId('pptx-peerless-editor')).toBeTruthy();
    expect(hybrid).not.toHaveBeenCalled();
    expect(engines).toHaveBeenCalledTimes(1);
    expect(warn).not.toHaveBeenCalled();
    expect(() => Reflect.get(api, 'handle')).toThrow('handleAsync');
    expect(() => Reflect.get(api, 'save')()).toThrow('saveAsync');
    expect(() => Reflect.get(api, 'handle')).toThrow('legacySyncApi');
    expect(() => Reflect.get(api, 'save')()).toThrow('legacySyncApi');
    expect(await api.saveAsync()).toBeInstanceOf(Uint8Array);
    expect((await story(api)).paragraphs).not.toHaveLength(0);
  } finally { hybrid.mockRestore(); engines.mockRestore(); warn.mockRestore(); }
});

for (const operation of ['handle', 'save', 'host'] as const) {
  test(`onReady ${operation} failures reach onError and the console once`, async () => {
    const initial = await open();
    initial.view.unmount();
    const diagnostic = spyOn(console, 'error').mockImplementation(() => {});
    const hostError = new Error('Host initialization failed');
    const onError = mock((_error: Error) => { throw new Error('Host error handler failed'); });
    const onReady = mock((api: PptxAsyncEditorApi) => {
      if (operation === 'handle') Reflect.get(api, 'handle');
      if (operation === 'save') Reflect.get(api, 'save')();
      throw hostError;
    });
    const view = render(<PptxEditor fonts={[]} file={initial.bytes} onReady={onReady} onError={onError} />);
    try {
      await waitFor(() => expect(onReady).toHaveBeenCalledTimes(1));
      await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
      const error = onError.mock.calls[0][0];
      if (operation === 'host') expect(error).toBe(hostError);
      else expect(error.message).toContain('legacySyncApi');
      expect(diagnostic.mock.calls.filter((call) => call.includes(error))).toHaveLength(1);
      await act(async () => { view.rerender(<PptxEditor fonts={[]} file={initial.bytes} onReady={onReady} onError={onError} />); });
      expect(onReady).toHaveBeenCalledTimes(1);
      expect(onError).toHaveBeenCalledTimes(1);
    } finally { view.unmount(); diagnostic.mockRestore(); }
  });
}

for (const capability of ['Worker', 'OffscreenCanvas'] as const) {
  test(`missing ${capability} keeps the fallback replica and provider stable when presence arrives`, async () => {
    const initial = await open();
    initial.view.unmount();
    const descriptor = Object.getOwnPropertyDescriptor(globalThis, capability);
    Object.defineProperty(globalThis, capability, { configurable: true, value: undefined });
    const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }];
    const publications: (AsyncCollaborationReplica | null)[] = [];
    const providers: CollaborationProvider[] = [];
    const apis: PptxAsyncEditorApi[] = [];
    const errors: Error[] = [];
    let connections = 0;
    let disconnections = 0;
    function Host() {
      const [replica, setReplica] = useState<AsyncCollaborationReplica | null>(null);
      const [presence, setPresence] = useState<CollaborationProvider>();
      const publish = useCallback((value: AsyncCollaborationReplica | null) => {
        publications.push(value);
        if (publications.length < 10) setReplica(value);
      }, []);
      useEffect(() => {
        if (!replica) return;
        const listeners = new Set<(event: CollaborationTransportEvent) => void>();
        const provider = new CollaborationProvider(replica, {
          connect() { connections++; for (const listener of listeners) listener({ type: 'open' }); },
          disconnect() { disconnections++; },
          send() { return true; },
          onEvent(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
        });
        providers.push(provider);
        setPresence(provider);
        provider.connect();
        return () => { provider.destroy(); setPresence(undefined); };
      }, [replica]);
      return <PptxEditor fonts={fonts} file={initial.bytes} collaboration={{ clientId: 700, onReplica: publish, presence }}
        onReady={(api) => { apis.push(api); }} onError={(error) => { errors.push(error); }} />;
    }
    const view = render(<Host />);
    try {
      await waitFor(() => expect(providers.length).toBeGreaterThan(0));
      await act(async () => { view.rerender(<Host />); await new Promise((resolve) => setTimeout(resolve, 20)); });
      expect(publications).toHaveLength(1);
      expect(providers).toHaveLength(1);
      expect(connections).toBe(1);
      expect(disconnections).toBe(0);
      expect(apis).toHaveLength(1);
      expect(errors).toEqual([]);
      expect(await publications[0]!.encodeStateVector()).toBeInstanceOf(Uint8Array);
      expect(await apis[0]!.saveAsync()).toBeInstanceOf(Uint8Array);
      view.unmount();
      expect(publications).toEqual([expect.anything(), null]);
      expect(disconnections).toBe(1);
    } finally {
      view.unmount();
      if (descriptor) Object.defineProperty(globalThis, capability, descriptor);
      else Reflect.deleteProperty(globalThis, capability);
    }
  });
}

test('in-thread room changes connect a desktop-style provider host without a disposed replica', async () => {
  const initial = await open();
  initial.view.unmount();
  const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }];
  const providers: { roomId: string; replica: AsyncCollaborationReplica; provider: CollaborationProvider }[] = [];
  const connections: string[] = [];
  const disconnections: string[] = [];
  const releases: Promise<unknown>[] = [];
  const errors: Error[] = [];
  function Host({ roomId }: { roomId: string }) {
    const [replica, setReplica] = useState<AsyncCollaborationReplica | null>(null);
    const [presence, setPresence] = useState<CollaborationProvider>();
    useEffect(() => {
      if (!replica) return;
      const listeners = new Set<(event: CollaborationTransportEvent) => void>();
      const provider = new CollaborationProvider(replica, {
        connect() { connections.push(roomId); for (const listener of listeners) listener({ type: 'open' }); },
        disconnect() { disconnections.push(roomId); },
        send() { return true; },
        onEvent(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
      });
      providers.push({ roomId, replica, provider });
      setPresence(provider);
      const unsubscribe = provider.onError((error) => { errors.push(error); });
      provider.connect();
      return () => {
        releases.push(replica.encodeStateVector().catch((error) => { errors.push(error); }));
        unsubscribe(); provider.destroy(); setPresence(undefined);
      };
    }, [replica, roomId]);
    return <PptxEditor experimentalWorkerOpen={false} fonts={fonts} file={initial.bytes}
      collaboration={{ roomId, clientId: 700, onReplica: setReplica, presence }}
      onError={(error) => { errors.push(error); }} />;
  }
  const view = render(<Host roomId="A" />);
  try {
    await waitFor(() => expect(connections).toEqual(['A']));
    const old = providers[0]!.replica;
    expect(() => view.rerender(<Host roomId="B" />)).not.toThrow();
    await waitFor(() => expect(providers[providers.length - 1]!.replica).not.toBe(old));
    const current = providers[providers.length - 1]!;
    expect(current.roomId).toBe('B');
    expect(current.provider.status).toBe('connected');
    expect(connections[connections.length - 1]).toBe('B');
    expect(disconnections).toContain('A');
    expect(errors).toEqual([]);
    expect(await current.replica.encodeStateVector()).toBeInstanceOf(Uint8Array);
    await expect(old.encodeStateVector()).rejects.toThrow();
    view.unmount();
    await Promise.all(releases);
    expect(errors).toEqual([]);
    expect(disconnections.filter((room) => room === 'B')).toHaveLength(connections.filter((room) => room === 'B').length);
  } finally { view.unmount(); }
});

for (const change of ['room', 'disconnect'] as const) {
  test(`in-thread collaboration ${change} disposes the old session and drops its edits`, async () => {
    const initial = await open();
    initial.view.unmount();
    const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }];
    const apis: PptxAsyncEditorApi[] = [];
    const publications: (AsyncCollaborationReplica | null)[] = [];
    const publish = (value: AsyncCollaborationReplica | null) => { publications.push(value); };
    const ui = (roomId?: string) => <PptxEditor experimentalWorkerOpen={false} fonts={fonts} file={initial.bytes} clientId={700}
      collaboration={roomId ? { roomId, clientId: 700, onReplica: publish } : undefined}
      onReady={(api) => { apis.push(api); }} />;
    const view = render(ui('first'));
    try {
      await waitFor(() => expect(apis).toHaveLength(1));
      const old = apis[0]!;
      const access = await old.handleAsync();
      const snapshot = await access.snapshot();
      await act(async () => { await access.setSlideNotes(snapshot.slides[0]!.id, 'ROOM_A'); });
      expect((await access.snapshot()).slides[0]!.notes).toBe('ROOM_A');
      view.rerender(ui(change === 'room' ? 'second' : undefined));
      await waitFor(() => expect(apis).toHaveLength(2));
      expect((await (await apis[1]!.handleAsync()).snapshot()).slides[0]!.notes).toBe(snapshot.slides[0]!.notes);
      await expect(old.saveAsync()).rejects.toThrow('no longer open');
      await expect(publications[0]!.encodeStateVector()).rejects.toThrow();
      expect(publications.slice(0, 2)).toEqual([expect.anything(), null]);
      if (change === 'room') expect(publications[2]).not.toBe(publications[0]);
      view.unmount();
      expect(publications[publications.length - 1]).toBeNull();
    } finally { view.unmount(); }
  });
}

test('initialSlide clamps on open and later changes preserve edits and the session', async () => {
  const opener = spyOn(peerlessSessionOpener, 'open');
  try {
    const options = { initialSlide: Number.MAX_SAFE_INTEGER };
    const { api, client, view, ui } = await open(options);
    const starts = opener.mock.calls.length;
    expect(client.state.frame?.slideIndex).toBe(1);
    const before = await story(api);
    const result = await api.applyEdits({ expectVersion: await api.version(), steps: [{
      op: 'insertText', target: { kind: 'range', slideId: before.slideId, shapeId: before.shapeId, storyId: before.storyId, start: 0, end: 0 }, at: 'start', text: 'kept ',
    }] });
    expect(result.ok).toBe(true);
    const version = await api.version();
    options.initialSlide = 1;
    await act(async () => { view.rerender(ui()); });
    expect(opener).toHaveBeenCalledTimes(starts);
    expect(client.state.frame?.slideIndex).toBe(1);
    expect(await api.version()).toBe(version);
    expect(JSON.stringify(await story(api))).toContain('kept ');
  } finally { opener.mockRestore(); }
});

test('a real editor keeps one collaboration session across replica publication and equivalent options', async () => {
  const initial = await open();
  initial.view.unmount();
  const opener = spyOn(peerlessSessionOpener, 'open');
  opener.mockClear();
  const publications: (AsyncCollaborationReplica | null)[] = [];
  function Host() {
    const [replica, setReplica] = useState<AsyncCollaborationReplica | null>(null);
    const publish = useCallback((value: AsyncCollaborationReplica | null) => {
      publications.push(value);
      if (publications.length < 7) setReplica(value);
    }, []);
    return <>
      <output data-testid="replica">{replica ? 'published' : 'waiting'}</output>
      <PptxEditor fonts={[]} file={initial.bytes} collaboration={{ clientId: 7, onReplica: publish }} />
    </>;
  }
  const view = render(<Host />);
  try {
    await waitFor(() => expect(publications.some(Boolean)).toBe(true));
    await act(async () => { view.rerender(<Host />); await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(opener).toHaveBeenCalledTimes(1);
    expect(view.getByTestId('replica').textContent).toBe('published');
    expect(publications).toHaveLength(1);
    expect(await publications[0]!.encodeStateVector()).toBeInstanceOf(Uint8Array);
    view.unmount();
    expect(publications).toEqual([expect.anything(), null]);
  } finally { view.unmount(); opener.mockRestore(); }
});

for (const capability of ['Worker', 'OffscreenCanvas', 'experimentalWorkerOpen'] as const) {
  test(`${capability === 'experimentalWorkerOpen' ? 'disabled' : 'missing'} ${capability} uses the in-thread async API and warns once`, async () => {
    const { view, bytes } = await open();
    view.unmount();
    resetPptxEngineWarningsForTests();
    const descriptor = Object.getOwnPropertyDescriptor(globalThis, capability);
    if (capability !== 'experimentalWorkerOpen') Object.defineProperty(globalThis, capability, { configurable: true, value: undefined });
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    const opener = spyOn(peerlessSessionOpener, 'open');
    opener.mockClear();
    const hybrid = spyOn(workerEditorSessionOpener, 'open');
    let api: PptxAsyncEditorApi | undefined;
    try {
      const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
      const props = { fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }], file: bytes,
        experimentalWorkerOpen: capability === 'experimentalWorkerOpen' ? false : undefined,
        onReady: (value: PptxAsyncEditorApi) => { api = value; } };
      const fallback = render(<PptxEditor {...props} />);
      await waitFor(() => expect(api).toBeDefined());
      expect(fallback.queryByTestId('pptx-peerless-editor')).toBeNull();
      expect(opener).not.toHaveBeenCalled();
      expect(hybrid).not.toHaveBeenCalled();
      expect('handle' in api!).toBe(false);
      expect('save' in api!).toBe(false);
      expect(await api!.saveAsync()).toBeInstanceOf(Uint8Array);
      expect(await api!.goToSlide(2)).toBe(true);
      await act(async () => { fallback.rerender(<PptxEditor {...props} readOnly />); });
      const access = await api!.handleAsync();
      await expect(access.deleteSlide('slide')).rejects.toMatchObject({ code: 'read-only' });
      fallback.unmount();
      api = undefined;
      render(<PptxEditor {...props} />);
      await waitFor(() => expect(api).toBeDefined());
      expect(warn).toHaveBeenCalledTimes(1);
      expect(warn.mock.calls[0]?.[0]).toContain(capability);
      expect(warn.mock.calls[0]?.[0]).toContain('in-thread');
    } finally {
      if (descriptor) Object.defineProperty(globalThis, capability, descriptor);
      else Reflect.deleteProperty(globalThis, capability);
      warn.mockRestore(); opener.mockRestore(); hybrid.mockRestore(); resetPptxEngineWarningsForTests();
    }
  });
}
