import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createDocxCommandConnection, DocxSourceWorkerClient, type DocxCommandCheckpoint } from '@betteroffice/docx/worker';
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
  const checkpoints: number[] = [];
  const phases: string[] = [];
  const replay: { kind: string; reason?: string }[] = [];
  const onPasteOptionNotRestored = mock(() => {});
  const connect = async (resume?: Resume) => {
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
    const connection = await createDocxCommandConnection(client, initial);
    const edit = connection.client.call.edit;
    connection.client.call.edit = async command => {
      const ack = await edit(command);
      if (command.replay) replay.push({ kind: command.kind, reason: ack.outcomes[0]?.reason });
      return ack;
    };
    return connection;
  };
  const hook = renderHook(() => usePeerlessEditSession({
    source: bytes, openConnection: () => connect(), canvasForPage: () => null, prepareGeometry: () => undefined,
    recovery: {
      reopen: connect,
      checkpoint: async resident => {
        const checkpoint = await resident.client.call.checkpoint();
        checkpoints.push(checkpoint.throughSeq);
        return checkpoint;
      },
      onRecovery: event => { phases.push(event.phase); },
    },
    onPasteOptionNotRestored,
  }));
  await waitFor(() => expect(hook.result.current.status).toBe('ready'), { timeout: 10_000 });
  const binding = hook.result.current.binding!;
  const prefix: ReturnType<typeof binding.session.submit>[] = [];
  for (let i = 0; i < 127; i++) {
    prefix.push(binding.session.submit('insertText', { text: 'i' }, { after: prefix.at(-1), fence: true }));
  }
  const outcomes = await Promise.all(prefix.map(ticket => ticket.ack));
  expect(outcomes.filter(outcome => outcome.status !== 'applied')).toEqual([]);
  return { binding, workers, checkpoints, phases, replay, onPasteOptionNotRestored };
}

const rich = { paragraphs: [{ text: 'Rich', runs: [{ text: 'Rich', format: { other: { bold: true } } }] }] };

for (const clear of ['edit', 'dismiss']) {
  test(`a due checkpoint waits through two Paste Options choices and runs on ${clear}`, async () => {
    const { binding, checkpoints } = await open();
    const { session, admission, viewStore } = binding;
    const recorder: { current: PeerlessPasteRecorder | null } = { current: null };
    const view = render(<PeerlessInput admission={admission} viewStore={viewStore} pasteRecorder={recorder}
      onPasteOptionsLiveChange={binding.setPasteOptionsLive}
      onReplacePasteFragment={(fragment, selection, variant) => submitInputIntent(admission, {
        kind: 'replacePasteFragment', payload: { fragment, selection,
          ...(variant === 'plain' ? { plain: true as const } : variant === 'merge' ? { merge: true as const } : {}) },
      }, { source: 'paste', fence: true })} />);
    const present = async (throughSeq: number) => {
      await act(async () => { await session.flush(); });
      expect(session.state.throughSeq).toBe(throughSeq);
      act(() => viewStore.installFrameGeometry({
        identity: viewStore.getSnapshot().identity, coversThroughSeq: throughSeq, frameEpoch: throughSeq, layoutRevision: 1,
        pages: [], selectionRects: [], caret: { x: 1, y: 1, width: 1, height: 14 },
      }));
      expect(checkpoints).toEqual([]);
    };
    act(() => recorder.current!(rich, submitPasteFragment(admission, rich, { source: 'paste', fence: true }, () => {})));
    await present(128);
    const control = await view.findByRole('button', { name: 'Paste Options (Ctrl)' });
    for (const [index, name] of ['Merge Formatting', 'Keep Text Only'].entries()) {
      fireEvent.click(control);
      fireEvent.click(view.getByRole('menuitemradio', { name }));
      await present(129 + index);
    }
    if (clear === 'edit') {
      fireEvent.input(view.getByRole('textbox'), { inputType: 'insertText', target: { value: '!' } });
    } else {
      fireEvent.pointerDown(document.body);
    }
    await waitFor(() => expect(checkpoints).toEqual([130]));
    await act(async () => { await session.flush(); });
    expect(session.state.throughSeq).toBe(clear === 'edit' ? 131 : 130);
    expect(view.queryByRole('button', { name: 'Paste Options (Ctrl)' })).toBeNull();
  }, 30_000);
}

test('an older checkpoint containing the paste keeps the original paste and replays dependent and later typing', async () => {
  const { binding: { session }, workers, checkpoints, phases, replay, onPasteOptionNotRestored } = await open();
  expect((await session.submit('pasteFragment', { fragment: rich }).ack).status).toBe('applied');
  await waitFor(() => expect(checkpoints).toEqual([128]));
  const choice = session.submit('replacePasteFragment', { fragment: { paragraphs: [{ text: 'Rich' }] }, plain: true });
  expect((await choice.ack).status).toBe('applied');
  expect((await session.submit('insertText', { text: 'After' }, { after: choice }).ack).status).toBe('applied');
  expect((await session.submit('insertText', { text: 'Opt' }).ack).status).toBe('applied');
  act(() => workers[0]!.onerror?.({ message: 'Worker stopped' } as ErrorEvent));
  await waitFor(() => expect(phases).toEqual(['recovering', 'recovered']), { timeout: 10_000 });
  expect(replay).toEqual([
    { kind: 'replacePasteFragment', reason: 'paste-changed' },
    { kind: 'insertText', reason: undefined }, { kind: 'insertText', reason: undefined },
  ]);
  expect(onPasteOptionNotRestored).toHaveBeenCalledTimes(1);
  const restored = workers[1]!.sessions[0]!;
  expect(restored.paragraphs('body')[0]!.text).toBe('Page 1' + 'i'.repeat(127) + 'RichAfterOpt');
  expect(restored.storySegments('body').find(segment => segment.kind === 'text' && segment.text.includes('Rich')))
    .toMatchObject({ attributes: { bold: true } });
}, 30_000);
