import { createCommandHost } from '../../../../shared/office-session/commands/host';
import { isCommandEnvelope } from '../../../../shared/office-session/commands/types';
import { isTransferResult, transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import type {
  DocxCommandApplyContext,
  DocxCommandAck,
  DocxCommandFramePayload,
  DocxCommandPayload,
  DocxCommandReply,
  DocxCommandSelection,
  DocxEditState,
  ResidentCommandHooks,
  ResidentCommandHost,
  ResidentCommandHostOptions,
  ResidentCommandInitReply,
} from './residentCommandProtocol';

export function createResidentCommandHost(options: ResidentCommandHostOptions): ResidentCommandHost {
  const { init, hooks, registry } = options;
  let version = init.version ?? 0;
  let throughSeq = init.throughSeq ?? 0;
  if (!Number.isSafeInteger(version) || version < 0) throw new TypeError('Invalid command version');
  let tail: Promise<void> = Promise.resolve();
  let frameTail: Promise<void> = Promise.resolve();
  let failure: unknown;
  let selection: DocxCommandSelection | null | undefined;
  let editState: DocxEditState | undefined;
  const selections = new Map<number, DocxCommandSelection | null>();
  const selectionLimit = Math.max(256, options.retainedOutcomes ?? 256);
  const shared = createCommandHost<
    DocxCommandPayload, DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload
  >({
    session: init.session,
    incarnation: init.incarnation,
    epoch: init.epoch,
    throughSeq,
    retainedOutcomes: options.retainedOutcomes,
    refusedSeqs: init.refusedSeqs,
    replayAppliedSeqs: init.replayAppliedSeqs,
    getVersion: () => version,
    async apply(command, context) {
      const applyContext: DocxCommandApplyContext = {
        ...context,
        ...(command.afterSeq !== undefined && selections.has(command.afterSeq)
          ? { predecessorSelection: structuredClone(selections.get(command.afterSeq)!) } : {}),
      };
      const preflight = await hooks.preflight(command, applyContext);
      const result = preflight.status === 'ready'
        ? await hooks.apply(preflight.prepared, applyContext) : preflight;
      if (result.status === 'applied') {
        if (version === Number.MAX_SAFE_INTEGER) throw new Error('Command version space is exhausted');
        version += 1;
      }
      const selection = hooks.selection();
      if (result.status === 'applied' || result.status === 'noop') {
        selections.set(command.seq, structuredClone(selection));
        if (selections.size > selectionLimit) selections.delete(selections.keys().next().value!);
      }
      return { ...result, selection, editState: hooks.editState() };
    },
    frame: (input) => renderFrame({ ...input, full: false }),
    snapshot: () => hooks.snapshot(),
  });

  function renderFrame(input: Parameters<ResidentCommandHooks['frame']>[0]):
    Promise<Awaited<ReturnType<ResidentCommandHooks['frame']>>> {
    const result = frameTail.then(() => hooks.frame(input));
    frameTail = result.then(() => {}, () => {});
    return result;
  }

  function serialize<T>(run: () => Promise<T>): Promise<T> {
    const result = tail.then(() => {
      if (failure !== undefined) throw failure;
      return run();
    });
    tail = result.then(() => {}, (error: unknown) => { failure = error; });
    return result;
  }

  function packState(reply: DocxCommandReply, includeState: boolean): DocxCommandReply {
    const ack: DocxCommandAck = { ...(isTransferResult(reply) ? reply.value : reply) };
    const nextSelection = hooks.selection();
    const nextEditState = hooks.editState();
    if (includeState || !sameCommandState(selection, nextSelection)) {
      selection = structuredClone(nextSelection);
      ack.selection = structuredClone(selection);
    } else delete ack.selection;
    if (includeState || !sameCommandState(editState, nextEditState)) {
      editState = structuredClone(nextEditState);
      ack.editState = structuredClone(editState);
    } else delete ack.editState;
    return isTransferResult(reply) ? transferable(ack, reply.transfer) : ack;
  }

  async function fullAck(): Promise<ResidentCommandInitReply | TransferResult<ResidentCommandInitReply>> {
    const rendered = await renderFrame({ coversThroughSeq: throughSeq, version, full: true });
    const frame = isTransferResult(rendered) ? rendered.value : rendered;
    if (frame.coversThroughSeq !== throughSeq || frame.baseRevision !== undefined) {
      throw new Error('Full frame does not cover the command boundary');
    }
    const ack: ResidentCommandInitReply = {
      session: init.session, incarnation: init.incarnation, epoch: init.epoch,
      version, throughSeq, outcomes: [], frame,
      selection: structuredClone(hooks.selection()), editState: structuredClone(hooks.editState()),
    };
    selection = structuredClone(ack.selection);
    editState = structuredClone(ack.editState);
    return isTransferResult(rendered) ? transferable(ack, rendered.transfer) : ack;
  }

  const host: ResidentCommandHost = {
    handler(context, command) {
      return serialize(async () => {
        if (isCommandEnvelope(command) && command.basis.epoch !== init.epoch) {
          return packState({
            session: init.session, incarnation: init.incarnation, epoch: init.epoch,
            version, throughSeq,
            outcomes: [{ seq: command.seq, status: 'refused', reason: 'stale-epoch', committed: false }],
          }, true);
        }
        const duplicate = command.seq <= throughSeq;
        const reply = await shared.handler(context, command);
        const ack = isTransferResult(reply) ? reply.value : reply;
        throughSeq = ack.throughSeq;
        return packState(reply, duplicate || ack.outcomes.some(({ status }) =>
          status === 'refused' || status === 'cancelled' || status === 'duplicate'));
      });
    },
    policy(command) {
      return isCommandEnvelope(command) && command.basis.epoch !== init.epoch
        ? { lane: 'input', mutates: true, reframes: true, reorderable: false }
        : shared.policy(command);
    },
    checkpoint: (context) => serialize(() => shared.checkpoint(context)),
    checkpointPolicy: shared.checkpointPolicy,
    frame: () => failure === undefined ? shared.frame() : Promise.reject(failure),
    withFrameWork(run) {
      const result = frameTail.then(run);
      frameTail = result.then(() => {}, () => {});
      return result;
    },
    initial: () => serialize(fullAck),
    requestFullFrame: () => serialize(fullAck),
  };
  registry.handlers.edit = host.handler;
  registry.handlers.checkpoint = host.checkpoint;
  registry.policies.edit = host.policy;
  registry.policies.checkpoint = host.checkpointPolicy;
  return host;
}

function sameCommandState(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object') return false;
  if (left instanceof Uint8Array || right instanceof Uint8Array) {
    return left instanceof Uint8Array && right instanceof Uint8Array && left.length === right.length &&
      left.every((value, index) => value === right[index]);
  }
  if (Array.isArray(left) !== Array.isArray(right)) return false;
  const a = left as Record<string, unknown>;
  const b = right as Record<string, unknown>;
  const keys = Object.keys(a);
  return keys.length === Object.keys(b).length && keys.every((key) =>
    Object.prototype.hasOwnProperty.call(b, key) && sameCommandState(a[key], b[key]));
}

export function packCommandAck(
  reply: DocxCommandReply,
  hooks: Pick<ResidentCommandHooks, 'selection' | 'editState'>,
  includeState: boolean
): DocxCommandReply {
  if (!includeState) return reply;
  const ack = isTransferResult(reply) ? reply.value : reply;
  const packed = { ...ack, selection: structuredClone(hooks.selection()), editState: structuredClone(hooks.editState()) };
  return isTransferResult(reply) ? transferable(packed, reply.transfer) : packed;
}
