import { expect, mock, test } from 'bun:test';
import { copyWorkerSelection, createClipboardPrefetch, writeClipboardCapture, isPasteFragmentAllowed } from './clipboard';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { createYrsSession, readClipboardSlice, type DocxClipboardFragment, type YrsStorySegment } from '@betteroffice/docx/yrs';

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

for (const cut of [false, true]) {
  for (const objectKind of ['chart', 'shape', 'image']) {
    for (const depth of [0, 2]) {
      for (const controlKind of [null, 'sdt', 'blockSdt']) {
        test(`${cut ? 'cut' : 'copy'} refuses a picture with ${objectKind} in a table at depth ${depth} through ${controlKind ?? 'no control'}`, async () => {
          const { binding, submit, readSelection } = fixture();
          const image = { src: 'data:image/png;base64,AA==' };
          const embed = (embedKind: string, payload: Record<string, unknown>): YrsStorySegment => ({ kind: 'embed', embedKind, payload, attributes: {} });
          const table = (story: string) => ({ grid: [2000], rows: [{ cells: [{ story, tcPr: {} }] }] });
          const stories: Record<string, YrsStorySegment[]> = { body: [embed('image', image), embed('table', table('cell0'))] };
          for (let level = 0; level < depth; level++) stories[`cell${level}`] = [embed('table', table(`cell${level + 1}`))];
          const object = embed(objectKind, image);
          stories[`cell${depth}`] = controlKind === 'sdt'
            ? [embed('sdt', { content: [{ kind: 'text', text: 'kept' }, { kind: objectKind, payload: image }] })]
            : controlKind === 'blockSdt' ? [embed('blockSdt', { story: 'control' })] : [object];
          stories.control = [{ kind: 'text', text: 'kept', attributes: {} }, object];
          const write = mock(async () => {});
          const writeSlice = mock(async () => {});
          await expect(copyWorkerSelection({ binding, readSelection, cut, isCurrent: () => true, write, writeSlice,
            readClipboardSlice: async () => readClipboardSlice({ storySegments: story => stories[story] ?? [], tablePayload: () => null },
              { story: 'body', from: 0, to: 2 }),
          })).rejects.toThrow('Picture selection contains unsupported clipboard objects');
          expect(write).not.toHaveBeenCalled();
          expect(writeSlice).not.toHaveBeenCalled();
          expect(submit).not.toHaveBeenCalled();
        });
      }
    }
  }
}

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

function sliceFixture(fragment: DocxClipboardFragment | null) {
  return { text: 'A\tB', fragment };
}

test('copy with a fragment writes both types through writeSlice', async () => {
  const { binding, readSelection } = fixture();
  const fragment: DocxClipboardFragment = {
    paragraphs: [{
      tables: [{
        tblPr: { width: 9360, widthType: 'dxa' },
        grid: [],
        rows: [{ trPr: {}, cells: [{ tcPr: {}, content: [{ kind: 'pilcrow' }] }] }],
      }],
      text: '',
    }],
  };
  const write = mock(async (_text: string) => {});
  const writeSlice = mock(async (_text: string, _html: string) => {});
  const readClipboardSlice = mock(async () => sliceFixture(fragment));
  expect(await copyWorkerSelection({ binding, readSelection, readClipboardSlice, write, writeSlice, isCurrent: () => true })).toBe(true);
  expect(write).not.toHaveBeenCalled();
  expect(writeSlice).toHaveBeenCalledTimes(1);
  expect(writeSlice.mock.calls[0]?.[0]).toBe('A\tB');
  expect(writeSlice.mock.calls[0]?.[1]).toContain('<table');
});

test('copy without a fragment writes plain text only', async () => {
  const { binding, readSelection } = fixture();
  const write = mock(async (_text: string) => {});
  const writeSlice = mock(async (_text: string, _html: string) => {});
  const readClipboardSlice = mock(async () => sliceFixture(null));
  expect(await copyWorkerSelection({ binding, readSelection, readClipboardSlice, write, writeSlice, isCurrent: () => true })).toBe(true);
  expect(write).toHaveBeenCalledWith('A\tB');
  expect(writeSlice).not.toHaveBeenCalled();
});

test('copy with empty slice text writes nothing', async () => {
  const { binding, readSelection } = fixture();
  const write = mock(async (_text: string) => {});
  const writeSlice = mock(async (_text: string, _html: string) => {});
  const readClipboardSlice = mock(async () => ({ text: '', fragment: null }));
  expect(await copyWorkerSelection({ binding, readSelection, readClipboardSlice, write, writeSlice, isCurrent: () => true })).toBe(false);
  expect(write).not.toHaveBeenCalled();
  expect(writeSlice).not.toHaveBeenCalled();
});

test('cut in cells mode copies but deletes nothing', async () => {
  const { binding, readSelection, submit } = fixture();
  (binding.viewStore.getSnapshot() as unknown as { editState: Record<string, unknown> }).editState = {
    readOnly: false,
    tableContext: { hasMultiCellSelection: true },
  };
  const writeSlice = mock(async (_text: string, _html: string) => {});
  const readClipboardSlice = mock(async () => sliceFixture(null));
  const write = mock(async (_text: string) => {});
  expect(await copyWorkerSelection({ binding, readSelection, readClipboardSlice, write, writeSlice, cut: true, isCurrent: () => true })).toBe(true);
  expect(write).toHaveBeenCalled();
  expect(submit).not.toHaveBeenCalled();
});

test('paste gating follows the insertTable command state', () => {
  expect(isPasteFragmentAllowed({ getState: () => ({ enabled: true }) })).toBe(true);
  expect(isPasteFragmentAllowed({ getState: () => ({ enabled: false }) })).toBe(false);
});


test('picture-only captures survive prefetch and copy or cut through every capture path', async () => {
  for (const cut of [false, true]) {
    for (const path of ['slice', 'capture', 'prefetched']) {
      const { binding, readSelection, selection, submit } = fixture();
      const fragment = { paragraphs: [{ text: '', inlines: [{ image: { src: 'data:image/png;base64,AA==', width: 120, height: 60 } }] }] };
      const slice = { text: '', fragment };
      binding.viewStore.subscribe = () => () => {};
      const prefetch = createClipboardPrefetch({ binding, readSelection, readClipboardSlice: async () => slice, isCurrent: () => true });
      try {
        const capture = await prefetch.read();
        expect(capture).not.toBeNull();
        expect(prefetch.take()).toEqual(capture);
        if (!capture) throw new Error('Missing picture capture');
        const values = new Map<string, string>();
        const data = { setData: (type: string, value: string) => { values.set(type, value); } } as unknown as DataTransfer;
        expect(writeClipboardCapture(data, capture)).toBe(true);
        expect(values.get('text/plain')).toBe('');
        expect(values.get('text/html')).toContain('<img');
        expect(values.get('text/html')).toContain('data-betteroffice-docx');
        const writeSlice = mock(async (_text: string, _html: string) => {});
        expect(await copyWorkerSelection({ binding, readSelection, cut, isCurrent: () => true,
          write: async () => { throw new Error('Expected rich clipboard'); }, writeSlice,
          ...(path === 'slice' ? { readClipboardSlice: async () => slice } :
            path === 'capture' ? { readCapture: async () => capture } : { data, prefetched: capture }),
        })).toBe(true);
        if (path !== 'prefetched') expect(writeSlice.mock.calls[0]?.[1]).toContain('<img');
        if (cut) expect(submit.mock.calls[0]?.[0]).toEqual({ kind: 'replaceSelection', payload: { text: '', selection } });
        else expect(submit).not.toHaveBeenCalled();
      } finally { prefetch.dispose(); }
    }
  }
});


test('failed rich picture clipboard writes never cut the selection', async () => {
  const { binding, submit, readSelection } = fixture();
  const fragment = { paragraphs: [{ text: '', inlines: [{ image: { src: 'data:image/png;base64,AA==' } }] }] };
  const write = mock(async (_text: string) => {});
  await expect(copyWorkerSelection({ binding, readSelection, cut: true, isCurrent: () => true, write,
    readClipboardSlice: async () => ({ text: '', fragment }),
    writeSlice: async () => { throw new Error('Rich clipboard denied'); },
  })).rejects.toThrow('Rich clipboard denied');
  expect(write).not.toHaveBeenCalled();
  expect(submit).not.toHaveBeenCalled();
});


test('table-only copy and cut fall back to plain text when rich writing fails or is unavailable', async () => {
  for (const cut of [false, true]) {
    for (const unavailable of [false, true]) {
      const { binding, submit, readSelection, selection } = fixture();
      const fragment: DocxClipboardFragment = { paragraphs: [{ text: '', tables: [{
        tblPr: {}, grid: [], rows: [{ trPr: {}, cells: [{ tcPr: {}, content: [{ kind: 'pilcrow' }] }] }],
      }] }] };
      const write = mock(async (_text: string) => {});
      const writeSlice = unavailable ? undefined : mock(async (_text: string, _html: string) => { throw new Error('Rich clipboard denied'); });
      expect(await copyWorkerSelection({ binding, readSelection, cut, isCurrent: () => true, write, writeSlice,
        readClipboardSlice: async () => sliceFixture(fragment),
      })).toBe(true);
      expect(write).toHaveBeenCalledWith('A\tB');
      if (cut) expect(submit.mock.calls[0]?.[0]).toEqual({ kind: 'replaceSelection', payload: { text: '', selection } });
      else expect(submit).not.toHaveBeenCalled();
    }
  }
});
