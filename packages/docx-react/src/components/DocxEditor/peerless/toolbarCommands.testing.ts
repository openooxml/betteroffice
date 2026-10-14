import { expect, mock, type Mock } from 'bun:test';
import type { DocxCommandSelection, YrsSelectionContext } from '@betteroffice/docx/yrs';
import { createPeerlessShellController } from './shellController';
import { createPeerlessViewStore } from './viewStore';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { PLAIN_CONTEXT } from '../../../commands/testing';
import { evaluateDocxCommand } from '../../../commands/evaluate';
import type { DocxCommandArgs, DocxCommandId } from '../../../commands/types';

type Submitted = { status: 'submitted'; ticket: { ack: Promise<{ status: 'applied' }> } };
type ToolbarFixture = {
  controller: ReturnType<typeof createPeerlessShellController>;
  binding: PeerlessEditSessionBinding;
  viewStore: ReturnType<typeof createPeerlessViewStore>;
  selection: DocxCommandSelection;
  format: Mock<(intent: unknown, options: unknown) => Submitted>;
  submit: Mock<(intent: unknown, options: unknown) => Submitted>;
  flush: Mock<() => Promise<void>>;
  read: Mock<() => Promise<{ status: string; version: string; value: { version: string; paragraphs: never[] } }>>;
  context(delta: Partial<YrsSelectionContext>): void;
  parity<K extends DocxCommandId>(id: K, args?: DocxCommandArgs[K]): void;
};

export function toolbarFixture(): ToolbarFixture {
  const selection = { anchor: { story: 'body', paraId: 'p1', offset: 0 }, focus: { story: 'body', paraId: 'p1', offset: 4 } };
  const viewStore = createPeerlessViewStore({ identity: { session: 'worker', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: selection, anchors: null,
    editState: { readOnly: false, toolbar: PLAIN_CONTEXT, effectiveFormatting: null, storedMarks: null,
      canUndo: true, canRedo: false, hyperlink: null, tableContext: null }, geometry: null });
  const applied = () => ({ status: 'submitted' as const, ticket: { ack: Promise.resolve({ status: 'applied' as const }) } });
  const format = mock((_intent: unknown, _options: unknown) => applied());
  const submit = mock((_intent: unknown, _options: unknown) => applied());
  const flush = mock(async () => {});
  const binding = { viewStore, admission: {}, session: { flush }, commands: { format, submit } } as unknown as PeerlessEditSessionBinding;
  const read = mock(async () => ({ status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } }));
  const controller = createPeerlessShellController({ read: read as never, focus() {}, zoom: 1 });
  controller.attach(binding);
  const context = (delta: Partial<typeof PLAIN_CONTEXT>) => {
    const view = viewStore.getSnapshot();
    viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection,
      editState: { ...view.editState!, toolbar: { ...PLAIN_CONTEXT, ...delta } } });
  };
  const parity = <K extends DocxCommandId>(id: K, args?: DocxCommandArgs[K]) => {
    const env = controller.environment();
    expect(controller.commands.store.getState(id, args)).toEqual(evaluateDocxCommand(id, args, { ...env, supportedCommands: undefined }));
  };
  return { controller, binding, viewStore, selection, format, submit, flush, read, context, parity };
}
