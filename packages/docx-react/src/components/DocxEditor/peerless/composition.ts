import type { CommandTicket } from '../../../../../../shared/office-session/commands/types';
import type {
  PeerlessAdmissionTarget,
  PeerlessInputAdmission,
} from './inputAdmission';

/** @internal */
export interface PeerlessCompositionState {
  id: number;
  slot: number;
  target: PeerlessAdmissionTarget;
  text: string;
  phase: 'active' | 'committed' | 'cancelled' | 'refused';
  ticket?: CommandTicket;
  recoverableText?: string;
}

/** @internal */
export type PeerlessCompositionEndResult =
  | { status: 'committed'; state: PeerlessCompositionState; ticket: CommandTicket }
  | { status: 'cancelled'; state: PeerlessCompositionState };

type CompositionRecord = {
  state: PeerlessCompositionState;
  result?: PeerlessCompositionEndResult;
};

const records = new WeakMap<PeerlessCompositionState, CompositionRecord>();

function recordFor(state: PeerlessCompositionState): CompositionRecord {
  let record = records.get(state);
  if (!record) {
    record = { state };
    records.set(state, record);
  }
  return record;
}

/** @internal */
export function beginComposition(
  admission: PeerlessInputAdmission,
  id: number
): PeerlessCompositionState {
  const { slot, target } = admission.reserveComposition();
  const state: PeerlessCompositionState = { id, slot, target, text: '', phase: 'active' };
  records.set(state, { state });
  return state;
}

/** @internal */
export function updateComposition(
  state: PeerlessCompositionState,
  text: string
): PeerlessCompositionState {
  const record = recordFor(state);
  if (record.state.phase !== 'active') return record.state;
  const updated = { ...record.state, text };
  record.state = updated;
  records.set(updated, record);
  return updated;
}

/** @internal */
export function endComposition(
  admission: PeerlessInputAdmission,
  state: PeerlessCompositionState,
  text: string
): PeerlessCompositionEndResult {
  const record = recordFor(state);
  if (record.result) return record.result;
  if (record.state.phase !== 'active') {
    const settled = record.state;
    record.result = settled.ticket
      ? { status: 'committed', state: settled, ticket: settled.ticket }
      : { status: 'cancelled', state: settled };
    return record.result;
  }

  const settled: PeerlessCompositionState = { ...record.state, text, phase: 'cancelled' };
  const admissionResult = admission.settleComposition(
    settled.slot,
    text.length ? { text, target: settled.target } : null
  );
  if (admissionResult.status === 'held') {
    throw new Error('composition settlement must release its slot');
  }
  record.state = settled;
  records.set(settled, record);
  if (admissionResult.status === 'ignored') {
    if (text.length) {
      settled.phase = 'refused';
      settled.recoverableText = text;
    }
    record.result = { status: 'cancelled', state: settled };
    return record.result;
  }

  const ticket = admissionResult.ticket;
  settled.phase = 'committed';
  settled.ticket = ticket;
  record.result = { status: 'committed', state: settled, ticket };
  const retainText = () => {
    settled.phase = 'refused';
    settled.recoverableText = settled.text;
  };
  void ticket.ack.then((outcome) => {
    if (outcome.status === 'refused' || outcome.status === 'cancelled' || outcome.committed === false) {
      retainText();
    }
  }, retainText);
  return record.result;
}
