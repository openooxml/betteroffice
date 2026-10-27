import { expect, mock, test } from 'bun:test';
import { copyWorkerSelection, createClipboardPrefetch } from './clipboard';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { DocxClipboardFragment, DocxClipboardSlice, DocxCommandSelection } from '@betteroffice/docx/yrs';

const range: DocxCommandSelection = { anchor: { story: 'body', paraId: 'a', offset: 1 }, focus: { story: 'body', paraId: 'a', offset: 4 } };
const fragment = { paragraphs: [{ text: '', tables: [{ tblPr: {}, grid: [], rows: [{ trPr: {}, cells: [{ tcPr: {}, content: [{ kind: 'text', text: 'bcd' }, { kind: 'pilcrow' }] }] }] }] }] } as DocxClipboardFragment;

function fixture(slice: DocxClipboardSlice = { text: 'bcd', fragment }) {
  const listeners = new Set<() => void>();
  let state = { identity: { session: 'clipboard', incarnation: 1, epoch: 1 }, version: 1, workerSelection: null as DocxCommandSelection | null, editState: { readOnly: false } };
  let tail: unknown;
  const submit = mock((_intent: unknown) => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }));
  const binding = {
    viewStore: { getSnapshot: () => state, subscribe: (listener: () => void) => { listeners.add(listener); return () => { listeners.delete(listener); }; } },
    admission: { getSnapshot: () => ({ tail, compositionSlot: null, heldIntents: 0 }) },
    commands: { submit }, session: { flush: async () => {} },
  } as unknown as PeerlessEditSessionBinding;
  const publish = (next: Partial<typeof state>) => { state = { ...state, ...next }; for (const listener of listeners) listener(); };
  const readClipboardSlice = mock(async (_selection: DocxCommandSelection) => slice);
  const readSelection = mock(async () => ({ paraId: 'a', selectedText: 'bcd', paragraphText: 'abcdef', before: 'a', after: 'ef' }));
  const prefetch = createClipboardPrefetch({ binding, isCurrent: () => true, readSelection, readClipboardSlice });
  const setPending = (value: unknown) => { tail = value; };
  return { binding, publish, readClipboardSlice, readSelection, prefetch, submit, setPending };
}

function transfer() {
  const data = new Map<string, string>();
  return { data, transfer: { setData: (type: string, value: string) => { data.set(type, value); }, getData: (type: string) => data.get(type) ?? '' } as unknown as DataTransfer };
}

async function settled(prefetch: ReturnType<typeof createClipboardPrefetch>) {
  for (let turn = 0; turn < 100 && !prefetch.take(); turn += 1) await new Promise((resolve) => setTimeout(resolve, 10));
  return prefetch.take();
}

const denied = async () => { throw new DOMException('Write permission denied.', 'NotAllowedError'); };

test('a settled range fills the copy event synchronously when the async clipboard is denied', async () => {
  const { binding, publish, readClipboardSlice, readSelection, prefetch, setPending } = fixture();
  setPending({ ack: Promise.resolve() });
  publish({ workerSelection: range });
  setPending(undefined);
  const capture = await settled(prefetch);
  expect(capture).not.toBeNull();
  expect(capture!.html).toContain('<table');
  Object.defineProperty(capture!.slice, 'fragment', { get: () => { throw new Error('Fragment accessed during copy'); } });
  expect(readClipboardSlice).toHaveBeenCalledTimes(1);
  const { data, transfer: clipboardData } = transfer();
  const copied = copyWorkerSelection({ binding, readSelection, readClipboardSlice, isCurrent: () => true,
    write: denied, writeSlice: denied, data: clipboardData, prefetched: prefetch.take() });
  expect(data.get('text/plain')).toBe('bcd');
  expect(data.get('text/html')).toContain('data-betteroffice-docx');
  expect(data.get('text/html')).toContain('<table');
  expect(await copied).toBe(true);
  expect(readClipboardSlice).toHaveBeenCalledTimes(1);
});

test('cut deletes the captured range after the event write and never after a failed write', async () => {
  const { binding, publish, readClipboardSlice, readSelection, prefetch, submit } = fixture();
  publish({ workerSelection: range });
  const prefetched = await settled(prefetch);
  const { data, transfer: clipboardData } = transfer();
  expect(await copyWorkerSelection({ binding, readSelection, readClipboardSlice, isCurrent: () => true, cut: true,
    write: denied, writeSlice: denied, data: clipboardData, prefetched })).toBe(true);
  expect(data.get('text/plain')).toBe('bcd');
  expect(submit.mock.calls[0]?.[0]).toEqual({ kind: 'replaceSelection', payload: { text: '', selection: range } });
  submit.mockClear();
  await expect(copyWorkerSelection({ binding, readSelection, readClipboardSlice, isCurrent: () => true, cut: true,
    write: denied, writeSlice: denied, data: transfer().transfer, prefetched: null })).rejects.toThrow('Write permission denied.');
  expect(submit).not.toHaveBeenCalled();
});

test('a stale capture is not taken and copy falls back to the async read', async () => {
  const { binding, publish, readClipboardSlice, readSelection, prefetch, setPending } = fixture();
  publish({ workerSelection: range });
  expect(await settled(prefetch)).not.toBeNull();
  setPending({ ack: Promise.resolve() });
  expect(prefetch.take()).toBeNull();
  setPending(undefined);
  publish({ version: 2 });
  expect(prefetch.take()).toBeNull();
  const { data, transfer: clipboardData } = transfer();
  const writeSlice = mock(async (_text: string, _html: string) => {});
  expect(await copyWorkerSelection({ binding, readSelection, readClipboardSlice, isCurrent: () => true,
    write: denied, writeSlice, data: clipboardData, prefetched: prefetch.take() })).toBe(true);
  expect(data.size).toBe(0);
  expect(writeSlice.mock.calls[0]?.[0]).toBe('bcd');
  expect(await settled(prefetch)).not.toBeNull();
  prefetch.dispose();
});

test('collapsed selections are never read ahead', async () => {
  const { publish, readClipboardSlice, prefetch } = fixture();
  publish({ workerSelection: { anchor: range.anchor, focus: range.anchor } });
  await new Promise((resolve) => setTimeout(resolve, 150));
  expect(readClipboardSlice).not.toHaveBeenCalled();
  expect(prefetch.take()).toBeNull();
  prefetch.dispose();
});

test('copying cells selected under a collapsed selection reads their slice on demand', async () => {
  const { binding, publish, readClipboardSlice, readSelection, prefetch } = fixture();
  publish({ workerSelection: { anchor: range.anchor, focus: range.anchor } });
  const writeSlice = mock(async (_text: string, _html: string) => {});
  expect(await copyWorkerSelection({ binding, readSelection, readClipboardSlice, isCurrent: () => true, write: denied, writeSlice,
    data: transfer().transfer, prefetched: prefetch.take(), readCapture: () => prefetch.read() })).toBe(true);
  expect(writeSlice.mock.calls[0]?.[0]).toBe('bcd');
  expect(writeSlice.mock.calls[0]?.[1]).toContain('<table');
  prefetch.dispose();
});


test('prefetch starts without a quiet timer and coalesces selection changes during a read', async () => {
  const f = fixture();
  let finish!: (slice: DocxClipboardSlice) => void;
  f.readClipboardSlice.mockImplementationOnce(() => new Promise(resolve => { finish = resolve; }));
  f.publish({ workerSelection: range });
  for (let i = 0; i < 20; i++) await Promise.resolve();
  expect(f.readClipboardSlice).toHaveBeenCalledTimes(1);
  for (let offset = 5; offset < 20; offset++) f.publish({ workerSelection: { ...range, focus: { ...range.focus, offset } } });
  expect(f.readClipboardSlice).toHaveBeenCalledTimes(1);
  finish({ text: 'old', fragment: null });
  for (let i = 0; i < 40; i++) await Promise.resolve();
  expect(f.readClipboardSlice).toHaveBeenCalledTimes(2);
  expect(f.prefetch.take()?.selection.focus.offset).toBe(19);
  f.prefetch.dispose();
});

for (const failed of [false, true]) test(`prefetch memoizes ${failed ? 'failed' : 'empty'} reads until the key changes`, async () => {
  const f = fixture({ text: '', fragment: null });
  if (failed) f.readClipboardSlice.mockRejectedValue(new Error('unavailable'));
  f.publish({ workerSelection: range });
  await new Promise(resolve => setTimeout(resolve, 350));
  expect(f.readClipboardSlice).toHaveBeenCalledTimes(1);
  expect(f.prefetch.take()).toBeNull();
  f.publish({ version: 2 });
  await new Promise(resolve => setTimeout(resolve, 150));
  expect(f.readClipboardSlice).toHaveBeenCalledTimes(2);
  f.prefetch.dispose();
});

test('an incarnation change invalidates a capture with the same version and selection', async () => {
  const f = fixture();
  f.publish({ workerSelection: range });
  expect(await settled(f.prefetch)).not.toBeNull();
  f.publish({ identity: { session: 'clipboard', incarnation: 2, epoch: 1 } });
  expect(f.prefetch.take()).toBeNull();
  expect(await settled(f.prefetch)).not.toBeNull();
  expect(f.readClipboardSlice).toHaveBeenCalledTimes(2);
  f.prefetch.dispose();
});


test('a pending cut shares the prefetch read and deletes only after the fresh event writes', async () => {
  const f = fixture();
  let finish!: (slice: DocxClipboardSlice) => void;
  f.readClipboardSlice.mockImplementation(() => new Promise(resolve => { finish = resolve; }));
  f.publish({ workerSelection: range });
  const { data, transfer: clipboardData } = transfer();
  const replay = mock((capture: NonNullable<ReturnType<typeof f.prefetch.take>>) => {
    expect(f.submit).not.toHaveBeenCalled();
    data.set('text/plain', capture.slice.text);
    return true;
  });
  const copied = copyWorkerSelection({ ...f, isCurrent: () => true, cut: true, data: clipboardData,
    readCapture: () => f.prefetch.read(), writeNative: replay, write: denied, writeSlice: denied });
  for (let i = 0; i < 20; i++) await Promise.resolve();
  expect(f.readClipboardSlice).toHaveBeenCalledTimes(1);
  expect(replay).not.toHaveBeenCalled();
  expect(f.submit).not.toHaveBeenCalled();
  finish({ text: 'bcd', fragment: null });
  expect(await copied).toBe(true);
  expect(data.get('text/plain')).toBe('bcd');
  expect(f.submit).toHaveBeenCalledTimes(1);
  f.prefetch.dispose();
});

test('a missed cut with unavailable activation and denied async writes preserves the document', async () => {
  const f = fixture();
  f.publish({ workerSelection: range });
  await expect(copyWorkerSelection({ ...f, isCurrent: () => true, cut: true,
    readCapture: () => f.prefetch.read(), writeNative: () => false, write: denied, writeSlice: denied })).rejects.toThrow('Write permission denied.');
  expect(f.submit).not.toHaveBeenCalled();
  f.prefetch.dispose();
});


test('an incarnation replaced during a pending cut cannot receive the old operation', async () => {
  const f = fixture();
  let finish!: (slice: DocxClipboardSlice) => void;
  f.readClipboardSlice.mockImplementationOnce(() => new Promise(resolve => { finish = resolve; }));
  f.publish({ workerSelection: range });
  const write = mock(async (_text: string) => {});
  const copied = copyWorkerSelection({ ...f, isCurrent: () => true, cut: true,
    readCapture: () => f.prefetch.read(), write });
  for (let i = 0; i < 20; i++) await Promise.resolve();
  f.publish({ identity: { session: 'clipboard', incarnation: 2, epoch: 1 } });
  finish({ text: 'old', fragment: null });
  expect(await copied).toBe(false);
  expect(write).not.toHaveBeenCalled();
  expect(f.submit).not.toHaveBeenCalled();
  f.prefetch.dispose();
});


test('a pending cut cannot adopt a selection changed while its read is in flight', async () => {
  const f = fixture();
  let finish!: (slice: DocxClipboardSlice) => void;
  f.readClipboardSlice.mockImplementationOnce(() => new Promise(resolve => { finish = resolve; }));
  f.publish({ workerSelection: range });
  const writeNative = mock(() => true);
  const copied = copyWorkerSelection({ ...f, isCurrent: () => true, cut: true,
    readCapture: () => f.prefetch.read(), writeNative, write: denied, writeSlice: denied });
  for (let i = 0; i < 20; i++) await Promise.resolve();
  f.publish({ workerSelection: { ...range, focus: { ...range.focus, offset: 5 } } });
  finish({ text: 'bcd', fragment: null });
  await expect(copied).rejects.toThrow('Selection read was superseded');
  expect(writeNative).not.toHaveBeenCalled();
  expect(f.submit).not.toHaveBeenCalled();
  f.prefetch.dispose();
});
