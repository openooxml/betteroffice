import { expect, mock, test } from 'bun:test';
import { copyWorkerSelection } from './clipboard';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { createYrsSession } from '@betteroffice/docx/yrs';

function fixture(readOnly = false) {
  const selection = { anchor: { story: 'body', paraId: 'a', offset: 1 }, focus: { story: 'body', paraId: 'a', offset: 4 } };
  const state = { workerSelection: selection, editState: { readOnly }, version: 1 };
  const submit = mock((_intent: unknown) => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }));
  const binding = { viewStore: { getSnapshot: () => state }, commands: { submit }, session: { flush: async () => {} } } as unknown as PeerlessEditSessionBinding;
  const readSelection = mock(async () => ({ paraId: 'a', selectedText: 'bcd', paragraphText: 'abcdef', before: 'a', after: 'ef' }));
  return { binding, submit, readSelection, state, selection };
}

test('cut deletes the captured range after clipboard completion even if the caret moved', async () => {
  const { binding, submit, readSelection, state, selection } = fixture();
  const write = mock(async (text: string) => {
    expect(text).toBe('bcd');
    state.workerSelection = { anchor: { ...selection.anchor, offset: 5 }, focus: { ...selection.focus, offset: 5 } };
  });
  expect(await copyWorkerSelection({ binding, readSelection, write, cut: true, isCurrent: () => true })).toBe(true);
  expect(submit.mock.calls[0]?.[0]).toEqual({ kind: 'replaceSelection', payload: { text: '', selection } });
});

test('failed clipboard writes and read-only cuts preserve the document', async () => {
  const { binding, submit, readSelection } = fixture();
  await expect(copyWorkerSelection({ binding, readSelection, cut: true, isCurrent: () => true,
    write: async () => { throw new Error('clipboard unavailable'); } })).rejects.toThrow('clipboard unavailable');
  expect(submit).not.toHaveBeenCalled();
  const locked = fixture(true);
  expect(await copyWorkerSelection({ ...locked, write: async () => {}, cut: true, isCurrent: () => true })).toBe(false);
  expect(locked.submit).not.toHaveBeenCalled();
});

test('replacement while copying cannot cut the new document', async () => {
  const { binding, submit, readSelection } = fixture();
  let current = true;
  expect(await copyWorkerSelection({ binding, readSelection, cut: true, isCurrent: () => current,
    write: async () => { current = false; } })).toBe(false);
  expect(submit).not.toHaveBeenCalled();
});

test('copy and cut project raw selection offsets across tracked deletions in the native engine', async () => {
  for (const cut of [false, true]) {
    const session = await createYrsSession();
    try {
      const { paraId } = session.createStory('body', 'XXabc');
      session.deleteRange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 2 } },
        { name: 'Writer', date: '2026-01-01T00:00:00Z' });
      const selected = { anchor: { story: 'body', paraId, offset: 2 }, focus: { story: 'body', paraId, offset: 3 } };
      const { binding, state } = fixture();
      state.workerSelection = selected;
      binding.commands.submit = (intent) => {
        if (intent.kind !== 'replaceSelection' || !intent.payload.selection) throw new Error('Expected captured cut');
        const { anchor, focus } = intent.payload.selection;
        session.deleteRange({ story: anchor.story, start: anchor, end: focus });
        return { status: 'submitted', echoEnabled: false,
          ticket: { id: 1, seq: 1, ack: Promise.resolve({ status: 'applied' }), presented: Promise.resolve('presented') } };
      };
      const write = mock(async (_text: string) => {});
      const options = { binding, cut, write, isCurrent: () => true,
        readSelection: async () => session.selectionText({ story: 'body', start: selected.anchor, end: selected.focus }) };
      expect(await copyWorkerSelection(options)).toBe(true);
      expect(write).toHaveBeenCalledWith('a');
      const read = session.readParagraphs({ view: 'accepted' });
      expect(read.ok && read.paragraphs[0]?.text).toBe(cut ? 'bc' : 'abc');
    } finally { session.destroy(); }
  }
});
