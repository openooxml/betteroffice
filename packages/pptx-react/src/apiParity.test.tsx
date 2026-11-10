import { expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { act, render, waitFor } from '@testing-library/react';
import type { AsyncCollaborationReplica } from '@betteroffice/pptx';
import { open, story } from './worker/peerlessTestHarness';
import { peerlessSessionOpener } from './worker/usePeerlessPresentation';
import { PptxEditor } from './PptxEditor';
import * as publicApi from './index';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';
import type { PptxPluginContext } from './plugins/types';

test('selectText results and selection direction match between in-thread and peerless editors',
  () => selectTextParity(false), 30_000);

test('plugin navigation selectText results and selection direction match between in-thread and peerless editors',
  () => selectTextParity(true), 30_000);

async function selectTextParity(viaPlugin: boolean) {
  let context: PptxPluginContext<null> | undefined;
  const plugin = publicApi.definePptxPlugin({
    id: 'select-text-parity',
    createState: () => null,
    initialize(next) { context = next; },
    onEvent(next) { context = next; },
  });
  const direction = () => {
    const target = context?.snapshot.selection?.target;
    return target?.kind === 'text' ? { anchor: target.anchor, focus: target.focus } : null;
  };
  const worker = await open({ plugins: [plugin] });
  const target = await story(worker.api);
  expect(target.text.length).toBeGreaterThanOrEqual(4);
  const insertion = 4;
  const inserted = '😀';
  const text = target.text.slice(0, insertion) + inserted + target.text.slice(insertion);
  const last = text.length;
  const outside = last + 1;
  const split = insertion + 1;
  const collapsed = { anchor: 2, focus: 2 };
  const cases = [
    { start: 1, end: 4, accepted: true, selection: { anchor: 1, focus: 4 } },
    { start: 4, end: 1, accepted: true, selection: { anchor: 4, focus: 1 } },
    { start: 2, end: 2, accepted: true, selection: collapsed },
    { start: last, end: last, accepted: true, selection: { anchor: last, focus: last } },
    { start: last, end: 1, accepted: true, selection: { anchor: last, focus: 1 } },
    { start: 1, end: last, accepted: true, selection: { anchor: 1, focus: last } },
    { start: -1, end: 1, accepted: false },
    { start: 1, end: -1, accepted: false },
    { start: outside, end: 1, accepted: false },
    { start: 1, end: outside, accepted: false },
    { start: outside, end: outside, accepted: false },
    { start: 1.5, end: 4, accepted: false },
    { start: 1, end: 4.5, accepted: false },
    { start: split, end: 1, accepted: false },
    { start: 1, end: split, accepted: false },
    { start: split, end: split, accepted: false },
  ];
  const exercise = async (api: PptxAsyncEditorApi) => {
    const current = await story(api);
    expect(current.text).toBe(target.text);
    await act(async () => {
      expect(await api.applyEdits({
        expectVersion: await api.version(),
        steps: [{ op: 'insertText', at: 'start', text: inserted, target: {
          kind: 'range', slideId: current.slideId, shapeId: current.shapeId, storyId: current.storyId,
          start: insertion, end: insertion,
        } }],
      })).toMatchObject({ ok: true, applied: true });
    });
    expect((await story(api)).text).toBe(text);
    await waitFor(() => expect(context).toBeDefined());
    const select = async (start: number, end: number): Promise<boolean> => {
      if (!viaPlugin) {
        return api.selectText({ slide: 1, shapeId: current.shapeId, storyId: current.storyId, start, end });
      }
      const result = await context!.navigation.selectText({
        slideId: current.slideId, shapeId: current.shapeId, storyId: current.storyId, anchor: start, focus: end,
      }, { expectVersion: await api.version() });
      return result.ok;
    };
    const offsets = [0];
    for (const character of text) offsets.push(offsets[offsets.length - 1]! + character.length);
    let refused = 0;
    const results = [];
    for (const { start, end, accepted, selection: selected } of cases) {
      const selection = selected ?? { anchor: 0, focus: offsets[++refused]! };
      const expected = {
        slideId: current.slideId, slide: 1,
        target: { kind: 'text' as const, shapeId: current.shapeId, storyId: current.storyId, ...selection },
      };
      if (!accepted) {
        await act(async () => { expect(await select(selection.anchor, selection.focus)).toBe(true); });
        await waitFor(() => expect(context?.snapshot.selection).toEqual(expected));
      }
      let result!: boolean;
      await act(async () => {
        result = await select(start, end);
      });
      expect(result).toBe(accepted);
      await waitFor(() => expect(context?.snapshot.selection).toEqual(expected));
      results.push({ accepted: result, selection: direction() });
    }
    return results;
  };

  const workerResults = await exercise(worker.api);
  worker.view.unmount();
  context = undefined;
  const font = await readFile(fileURLToPath(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
  let api: PptxAsyncEditorApi | undefined;
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const inThread = render(<PptxEditor file={worker.bytes} fonts={[{ family: 'Liberation Sans', bytes: new Uint8Array(font) }]}
    experimentalWorkerOpen={false} plugins={[plugin]} onReady={(ready) => { api = ready; }} />);
  try {
    await waitFor(() => expect(api).toBeDefined(), { timeout: 15_000 });
    expect(await exercise(api!)).toEqual(workerResults);
  } finally { inThread.unmount(); warn.mockRestore(); }
}

for (const inThread of [false, true]) {
  for (const snapshot of [false, true]) {
    test(`API edits notify once per version with inThread=${inThread} snapshot=${snapshot}`, async () => {
      const changed = mock((_change: { version: string }) => {});
      const snapshots = mock(() => {});
      const callbacks = { onDocumentChange: changed, onChange: snapshot ? snapshots : undefined };
      const value = await open(inThread ? {} : callbacks);
      let view: Pick<typeof value.view, 'unmount'> = value.view;
      let api: PptxAsyncEditorApi | undefined = inThread ? undefined : value.api;
      const warn = spyOn(console, 'warn').mockImplementation(() => {});
      try {
        if (inThread) {
          view.unmount();
          const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
          view = render(<PptxEditor fonts={[{ family: 'Liberation Sans', bytes: new Uint8Array(font) }]} file={value.bytes} experimentalWorkerOpen={false}
            {...callbacks} onReady={(ready) => { api = ready; }} />);
          await waitFor(() => expect(api).toBeDefined());
        }
        const access = await api!.handleAsync();
        const slideId = (await access.snapshot()).slides[0]!.id;
        const before = await api!.version();
        expect(changed).not.toHaveBeenCalled();
        expect(snapshots).not.toHaveBeenCalled();
        await act(async () => { await access.setSlideNotes(slideId, 'API speaker notes'); });
        const version = await api!.version();
        expect(version).not.toBe(before);
        await waitFor(() => expect(changed).toHaveBeenCalledTimes(1));
        expect(changed.mock.calls[0]).toEqual([{ version }]);
        expect(snapshots).toHaveBeenCalledTimes(snapshot ? 1 : 0);
        await act(async () => { await api!.refresh(); await api!.goToSlide(1); });
        expect(changed).toHaveBeenCalledTimes(1);
        expect(snapshots).toHaveBeenCalledTimes(snapshot ? 1 : 0);
      } finally { view.unmount(); warn.mockRestore(); }
    });
  }
}

test('the root exports the command admission error and no unavailable error', () => {
  const error = new publicApi.PptxCommandAdmissionError('read-only');
  expect(error).toBeInstanceOf(Error);
  expect(error).toMatchObject({ name: 'PptxCommandAdmissionError', code: 'read-only' });
  expect(publicApi).not.toHaveProperty('PptxAsyncEditorUnavailableError');
});

test('peer-less collaboration without a clientId generates one for the session and the replica', async () => {
  const initial = await open();
  initial.view.unmount();
  const opener = spyOn(peerlessSessionOpener, 'open');
  opener.mockClear();
  const replicas: (AsyncCollaborationReplica | null)[] = [];
  const onReplica = (replica: AsyncCollaborationReplica | null) => { replicas.push(replica); };
  const view = render(<PptxEditor fonts={[]} file={initial.bytes} collaboration={{ onReplica }} />);
  try {
    await waitFor(() => expect(replicas.some(Boolean)).toBe(true));
    await act(async () => {
      view.rerender(<PptxEditor fonts={[]} file={initial.bytes} collaboration={{ onReplica }} />);
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
    expect(opener).toHaveBeenCalledTimes(1);
    const clientId = replicas[0]!.clientId;
    expect(Number.isSafeInteger(clientId)).toBe(true);
    expect(opener.mock.calls[0]![1]).toMatchObject({ collaboration: { clientId } });
    expect(await replicas[0]!.encodeStateVector()).toBeInstanceOf(Uint8Array);
  } finally { view.unmount(); opener.mockRestore(); }
});

test('in-thread collaboration without a clientId publishes a replica', async () => {
  const { view, bytes } = await open();
  view.unmount();
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const replicas: (AsyncCollaborationReplica | null)[] = [];
  const fallback = render(<PptxEditor fonts={[]} file={bytes} experimentalWorkerOpen={false}
    collaboration={{ onReplica: (replica) => { replicas.push(replica); } }} />);
  try {
    await waitFor(() => expect(replicas.some(Boolean)).toBe(true));
    expect(Number.isSafeInteger(replicas.find(Boolean)!.clientId)).toBe(true);
  } finally { fallback.unmount(); warn.mockRestore(); }
});

test('in-thread collaboration generates a new clientId when the file is replaced', async () => {
  const { view, bytes } = await open();
  view.unmount();
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const replicas: AsyncCollaborationReplica[] = [];
  const onReplica = (replica: AsyncCollaborationReplica | null) => { if (replica) replicas.push(replica); };
  const fallback = render(<PptxEditor fonts={[]} file={bytes} experimentalWorkerOpen={false} collaboration={{ onReplica }} />);
  try {
    await waitFor(() => expect(replicas).toHaveLength(1));
    await act(async () => {
      fallback.rerender(<PptxEditor fonts={[]} file={bytes.slice()} experimentalWorkerOpen={false} collaboration={{ onReplica }} />);
    });
    await waitFor(() => expect(replicas).toHaveLength(2));
    expect(replicas[1]!.clientId).not.toBe(replicas[0]!.clientId);
  } finally { fallback.unmount(); warn.mockRestore(); }
});

test('in-thread collaboration generates a new clientId when fonts reopen the deck', async () => {
  const { view, bytes } = await open();
  view.unmount();
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const replicas: AsyncCollaborationReplica[] = [];
  const onReplica = (replica: AsyncCollaborationReplica | null) => { if (replica) replicas.push(replica); };
  const fallback = render(<PptxEditor fonts={[]} file={bytes} experimentalWorkerOpen={false} collaboration={{ onReplica }} />);
  try {
    await waitFor(() => expect(replicas).toHaveLength(1));
    await act(async () => {
      fallback.rerender(<PptxEditor fonts={[{ family: 'Liberation Sans', bytes: new Uint8Array(font) }]} file={bytes}
        experimentalWorkerOpen={false} collaboration={{ onReplica }} />);
    });
    await waitFor(() => expect(replicas).toHaveLength(2));
    expect(replicas[1]!.clientId).not.toBe(replicas[0]!.clientId);
  } finally { fallback.unmount(); warn.mockRestore(); }
});

test('the peer-less async API exports structured content and markdown like the handle', async () => {
  const value = await open();
  try {
    const access = await value.api.handleAsync();
    const structured = await value.api.exportStructured();
    expect(structured.ok).toBe(true);
    expect(structured).toEqual(await access.exportStructured());
    const markdown = await value.api.exportMarkdown({ includeNotes: false });
    expect(markdown.ok).toBe(true);
    expect(markdown).toEqual(await access.exportMarkdown({ includeNotes: false }));
  } finally { value.view.unmount(); }
});

test('the in-thread async API exports structured content and markdown like the handle', async () => {
  const { view, bytes } = await open();
  view.unmount();
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  let api: PptxAsyncEditorApi | undefined;
  const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const fallback = render(<PptxEditor fonts={[{ family: 'Liberation Sans', bytes: new Uint8Array(font) }]} file={bytes}
    experimentalWorkerOpen={false} onReady={(value) => { api = value; }} />);
  try {
    await waitFor(() => expect(api).toBeDefined());
    const access = await api!.handleAsync();
    const structured = await api!.exportStructured();
    expect(structured.ok).toBe(true);
    expect(structured).toEqual(await access.exportStructured());
    const markdown = await api!.exportMarkdown();
    expect(markdown.ok).toBe(true);
    expect(markdown).toEqual(await access.exportMarkdown());
    expect((await api!.save()).byteLength).toBeGreaterThan(0);
    expect((await api!.saveAsync()).byteLength).toBeGreaterThan(0);
  } finally { fallback.unmount(); warn.mockRestore(); }
});
