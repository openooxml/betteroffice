import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { act, render, waitFor } from '@testing-library/react';
import type { AsyncCollaborationReplica } from '@betteroffice/pptx';
import { open } from './worker/peerlessTestHarness';
import { peerlessSessionOpener } from './worker/usePeerlessPresentation';
import { PptxEditor } from './PptxEditor';
import * as publicApi from './index';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';

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
  } finally { fallback.unmount(); warn.mockRestore(); }
});
