import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createDocxCommandConnection, DocxSourceWorkerClient, type DocxCommandCheckpoint, type DocxCommandConnection } from '@betteroffice/docx/worker';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { preloadEditWasm } from '@betteroffice/docx/yrs/wasm/index';
import { preloadOpcWasm } from '@betteroffice/docx/wasm/opc';
import { pagedDocx } from '../__fixtures__/pagedDocx';
import { PeerlessInput, type PeerlessPasteRecorder } from '../peerless/PeerlessInput';
import { submitInputIntent, submitPasteFragment } from '../peerless/inputAdmission';
import { usePeerlessEditSession, type UsePeerlessEditSessionOptions } from './usePeerlessEditSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, renderHook, waitFor } = await import('@testing-library/react');
await import('../pasteOptions/PasteOptionsControl');
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  const generated = resolve(import.meta.dir, '../../../../../docx/src/wasm/generated');
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(generated, 'edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(generated, 'opc/ooxml_opc_bg.wasm'))));
  startWorker = await residentWorkerFactory();
}, 30_000);
afterEach(cleanup);
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

type Resume = Parameters<NonNullable<UsePeerlessEditSessionOptions['recovery']>['reopen']>[0];

async function open() {
  const bytes = new Uint8Array(await pagedDocx(1));
  const workers: InProcessResidentWorker[] = [];
  const checkpoints: ReturnType<DocxCommandConnection['client']['call']['checkpoint']>[] = [];
  const events: string[] = [];
  const phases: string[] = [];
  const resumes: Resume[] = [];
  const connect = async (resume?: Resume) => {
    if (resume) resumes.push(resume);
    const worker = startWorker(7140);
    workers.push(worker);
    const client = new DocxSourceWorkerClient(worker);
    const checkpoint = resume?.checkpoint as DocxCommandCheckpoint | undefined;
    await client.openSource({ bytes, state: checkpoint?.state });
    const paragraph = worker.sessions[0]!.paragraphs('body')[0]!;
    const end = { story: 'body', paraId: paragraph.paraId, offset: paragraph.text.length };
    const initial = await client.commandInit({
      session: 'paste-recovery', epoch: 1, incarnation: resume?.incarnation ?? 1,
      version: checkpoint?.version, throughSeq: checkpoint?.throughSeq, recoveryState: checkpoint?.recoveryState,
      replayAppliedSeqs: resume?.replayAppliedSeqs, refusedSeqs: resume?.refusedSeqs,
      selection: { anchor: end, focus: end }, readOnly: false,
      fonts: [new Uint8Array(readFileSync(resolve(import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')))],
      layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
        measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} }),
      layoutWithRegions: true, layoutRevision: 1, extras: '{}',
    });
    return createDocxCommandConnection(client, initial);
  };
  const hook = renderHook(() => usePeerlessEditSession({
    source: bytes, openConnection: () => connect(), canvasForPage: () => null, prepareGeometry: () => undefined,
    recovery: {
      reopen: connect,
      checkpoint: resident => {
        events.push('checkpoint');
        const rpc = resident.client.call.checkpoint();
        checkpoints.push(rpc);
        return rpc;
      },
      onRecovery: event => { phases.push(event.phase); },
    },
  }));
  await waitFor(() => expect(hook.result.current.status).toBe('ready'), { timeout: 10_000 });
  const binding = hook.result.current.binding!;
  const prefix: ReturnType<typeof binding.session.submit>[] = [];
  for (let i = 0; i < 127; i++) {
    prefix.push(binding.session.submit('insertText', { text: 'i' }, { after: prefix.at(-1), fence: true }));
  }
  const outcomes = await Promise.all(prefix.map(ticket => ticket.ack));
  expect(outcomes.filter(outcome => outcome.status !== 'applied')).toEqual([]);
  const { session, admission, viewStore } = binding;
  const recorder: { current: PeerlessPasteRecorder | null } = { current: null };
  const view = render(<PeerlessInput admission={admission} viewStore={viewStore} pasteRecorder={recorder}
    onPasteOptionsLiveChange={dismiss => {
      events.push(dismiss ? 'live' : 'dismiss');
      binding.setPasteOptionsLive!(dismiss);
    }}
    onReplacePasteFragment={(fragment, selection, variant) => submitInputIntent(admission, {
      kind: 'replacePasteFragment', payload: { fragment, selection,
        ...(variant === 'plain' ? { plain: true as const } : variant === 'merge' ? { merge: true as const } : {}) },
    }, { source: 'paste', fence: true })} />);
  const present = async (throughSeq: number) => {
    await act(async () => { await session.flush(); });
    expect(session.state.throughSeq).toBe(throughSeq);
    expect(session.state.version).toBe(throughSeq);
    act(() => viewStore.installFrameGeometry({
      identity: viewStore.getSnapshot().identity, coversThroughSeq: throughSeq, frameEpoch: throughSeq, layoutRevision: 1,
      pages: [], selectionRects: [], caret: { x: 1, y: 1, width: 1, height: 14 },
    }));
  };
  const rich = { paragraphs: [{ text: 'Rich', runs: [{ text: 'Rich', format: { other: { bold: true } } }] }] };
  const paste = () => act(() => recorder.current!(rich, submitPasteFragment(admission, rich, { source: 'paste', fence: true }, () => {})));
  const control = { name: 'Paste Options (Ctrl)' };
  const choose = (name: string) => {
    fireEvent.click(view.getByRole('button', control));
    fireEvent.click(view.getByRole('menuitemradio', { name }));
  };
  const edit = () => fireEvent.input(view.getByRole('textbox'), { inputType: 'insertText', target: { value: '!' } });
  const dismiss = () => fireEvent.pointerDown(document.body);
  const checkpoint = async () => {
    await waitFor(() => expect(checkpoints).toHaveLength(1));
    return await checkpoints[0]!;
  };
  paste();
  await present(128);
  await view.findByRole('button', control);
  expect(checkpoints).toHaveLength(0);
  return { session, view, control, checkpoint, checkpoints, choose, present, paste, edit, dismiss, phases, workers, events, resumes };
}

for (const clear of ['edit', 'dismiss'] as const) {
  test(`a due checkpoint waits through both choices and runs on ${clear}`, async () => {
    const editor = await open();
    editor.choose('Merge Formatting');
    await editor.present(129);
    editor.choose('Keep Text Only');
    await editor.present(130);
    expect(editor.checkpoints).toHaveLength(0);
    editor[clear]();
    expect((await editor.checkpoint()).throughSeq).toBe(130);
    await editor.present(clear === 'edit' ? 131 : 130);
    expect(editor.view.queryByRole('button', editor.control)).toBeNull();
  }, 30_000);
}

test('dismissing a submitted choice waits for its acknowledgement before checkpointing', async () => {
  const editor = await open();
  const worker = editor.workers[0]!;
  const edits = () => worker.requests.filter(type => type === 'edit').length;
  const sent = edits();
  worker.hold();
  fireEvent.click(editor.view.getByRole('button', editor.control));
  act(() => {
    fireEvent.click(editor.view.getByRole('menuitemradio', { name: 'Keep Text Only' }));
    editor.dismiss();
  });
  await waitFor(() => expect(edits()).toBe(sent + 1));
  expect(editor.checkpoints).toHaveLength(0);
  expect(editor.session.state.throughSeq).toBe(128);
  worker.release();
  expect((await editor.checkpoint()).throughSeq).toBe(129);
}, 30_000);

test('the cap dismisses Paste Options before checkpointing and recovery restores all edits', async () => {
  const editor = await open();
  for (let seq = 129; seq <= 159; seq++) {
    editor.paste();
    await editor.present(seq);
    expect(editor.view.getByRole('button', editor.control)).toBeTruthy();
    expect(editor.checkpoints).toHaveLength(0);
  }
  editor.events.length = 0;
  editor.choose('Keep Text Only');
  await editor.present(160);
  expect((await editor.checkpoint()).throughSeq).toBe(160);
  expect(editor.events).toEqual(['dismiss', 'checkpoint']);
  expect(editor.view.queryByRole('button', editor.control)).toBeNull();
  expect(editor.view.queryByRole('menuitemradio')).toBeNull();
  editor.edit();
  await editor.present(161);
  act(() => editor.workers[0]!.onerror?.({ message: 'Worker stopped' } as ErrorEvent));
  await waitFor(() => expect(editor.phases).toEqual(['recovering', 'recovered']), { timeout: 10_000 });
  expect(editor.resumes.map(resume => resume.checkpoint?.throughSeq)).toEqual([160]);
  const restored = editor.workers[1]!.sessions[0]!;
  expect(restored.paragraphs('body')[0]!.text).toBe('Page 1' + 'i'.repeat(127) + 'Rich'.repeat(32) + '!');
}, 30_000);
