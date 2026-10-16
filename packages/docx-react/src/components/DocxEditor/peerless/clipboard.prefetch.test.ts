import { expect, mock, test } from 'bun:test';
import { copyWorkerSelection, createClipboardPrefetch } from './clipboard';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { DocxClipboardFragment, DocxClipboardSlice, DocxCommandSelection } from '@betteroffice/docx/yrs';

const range: DocxCommandSelection = { anchor: { story: 'body', paraId: 'a', offset: 1 }, focus: { story: 'body', paraId: 'a', offset: 4 } };
const fragment = { paragraphs: [{ text: 'bcd', segments: [] }] } as unknown as DocxClipboardFragment;

function fixture(slice: DocxClipboardSlice = { text: 'bcd', fragment }) {
  const listeners = new Set<() => void>();
  let state = { version: 1, workerSelection: null as DocxCommandSelection | null, editState: { readOnly: false } };
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
  expect(await settled(prefetch)).not.toBeNull();
  expect(readClipboardSlice).toHaveBeenCalledTimes(1);
  const { data, transfer: clipboardData } = transfer();
  const copied = copyWorkerSelection({ binding, readSelection, readClipboardSlice, isCurrent: () => true,
    write: denied, writeSlice: denied, data: clipboardData, prefetched: prefetch.take() });
  expect(data.get('text/plain')).toBe('bcd');
  expect(data.get('text/html')).toContain('data-betteroffice-docx');
  expect(data.get('text/html')).toContain('<p>bcd</p>');
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
