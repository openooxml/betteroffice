import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import type { DocxCommandAck, DocxCommandFrame, ResidentCommandConnection } from '@betteroffice/docx/peerless';
import { usePeerlessEditSession, type UsePeerlessEditSessionOptions } from './usePeerlessEditSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

function stagedFrame(): DocxCommandFrame {
  const bitmap = () => ({ width: 1, height: 1, close: mock(() => {}) }) as unknown as ImageBitmap;
  return { coversThroughSeq: 0, revision: 0, payload: {
    frame: new ArrayBuffer(0), layoutRevision: 1, caret: { frameEpoch: 0, caretRect: null },
    stagedPresentation: { frameEpoch: 0, pages: [{
      page: { pageId: '1', pageIndex: 0 }, frameEpoch: 0, bitmap: bitmap(),
      backgrounds: [{ id: 'background', rect: { x: 0, y: 0, width: 1, height: 1 }, bitmap: bitmap() }],
    }] },
  } };
}

function resident() {
  const initial: ResidentCommandConnection['initial'] = {
    session: 'lifecycle', incarnation: 1, epoch: 1, version: 0, throughSeq: 0, outcomes: [], selection: null,
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: null }, frame: stagedFrame(),
  };
  const dispose = mock(async () => {});
  const connection: ResidentCommandConnection = {
    initial, requestFullFrame: async () => initial,
    client: { failure: undefined, onFailure: () => () => {}, dispose,
      call: { edit: async () => initial,
        checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }) },
    },
  };
  const options: UsePeerlessEditSessionOptions = {
    source: Uint8Array.of(1), openConnection: async () => connection,
    canvasForPage: () => null, prepareGeometry: () => undefined,
  };
  return { connection, dispose, options };
}

function expectClosed(frame: DocxCommandFrame) {
  const page = frame.payload.stagedPresentation!.pages[0]!;
  expect(page.bitmap.close).toHaveBeenCalledTimes(1);
  expect(page.backgrounds![0]!.bitmap.close).toHaveBeenCalledTimes(1);
}

test('unmounting during connection open releases the unused initial frame', async () => {
  const { connection, dispose, options } = resident();
  let resolve!: (value: ResidentCommandConnection) => void;
  options.openConnection = () => new Promise((yes) => { resolve = yes; });
  const view = renderHook(() => usePeerlessEditSession(options));
  view.unmount();
  await act(async () => { resolve(connection); });
  expectClosed(connection.initial.frame);
  expect(dispose).toHaveBeenCalledTimes(1);
});

test('a requested frame arriving after unmount releases its transferred bitmaps', async () => {
  const { connection, options } = resident();
  let resolve!: (value: DocxCommandAck) => void;
  const request = mock(() => new Promise<DocxCommandAck>((yes) => { resolve = yes; }));
  connection.requestFullFrame = request;
  const view = renderHook(() => usePeerlessEditSession(options));
  await waitFor(() => expect(view.result.current.status).toBe('ready'));
  expect(request).toHaveBeenCalledTimes(1);
  view.unmount();
  const frame = stagedFrame();
  await act(async () => { resolve({ ...connection.initial, frame }); });
  expectClosed(frame);
});
