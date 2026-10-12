import { createWorkbookCommandState, projectEditState, resolveCommandTarget, type WorkbookCommandState } from './commandHandlers';
import type { CommandIdentity, CommandStatus, WorkbookCommandSelection } from './commandProtocol';
import type { CapturedFormat, WorkbookCalculationContext, WorkbookHandle } from '../wasm/loader';
import type { CommandOutcome } from '../../../../shared/office-session/commands/types';

export interface WorkbookRecoveryReplaySeed {
  capture(state: WorkbookCommandState): unknown;
  restore(handle: WorkbookHandle): void;
}

const noReplaySeed: WorkbookRecoveryReplaySeed = { capture: () => undefined, restore: () => {} };

interface RecoveryMetadata {
  session: string;
  epoch: number;
  throughSeq: number;
  revision: number;
  engineVersion: string;
  versions: [number, string][];
  bases: [string, number][];
  refusedSeqs: number[];
  refusedRanges: [number, number][];
  refusedOutcomes: [number, CommandOutcome][];
  statuses: [number, CommandStatus][];
  selection: WorkbookCommandSelection;
  readOnly: boolean;
  sheetIds: string[];
  nextSheetId: number;
  sheetRevision: number;
  calculation: WorkbookCalculationContext | null;
  formats: [string, CapturedFormat][];
  hostVersion: number;
  dirty: boolean;
  historySelections?: WorkbookCommandState['historySelections'];
}

export interface WorkbookRecovery extends RecoveryMetadata {
  engine: Uint8Array;
}

const MAGIC = new Uint8Array([88, 76, 83, 88, 82, 69, 67, 86]);
const CODEC_VERSION = 3;
const HEADER_BYTES = 21;

export function isCollaborativeWorkbookRecovery(engine: Uint8Array): boolean {
  const unsupported = (): never => { throw new TypeError('Unsupported workbook recovery checkpoint'); };
  if (engine.length < 42 || !MAGIC.every((byte, index) => engine[index] === byte) || engine[8] !== 3) return unsupported();
  // Engine metadata follows the codec header and SHA-256 digest.
  let offset = 41;
  let length = 0;
  for (let shift = 0; shift <= 28; shift += 7) {
    const byte = engine[offset++];
    if (byte === undefined) return unsupported();
    length += (byte & 0x7f) * 2 ** shift;
    if (byte < 0x80) {
      if (length > engine.length - offset) return unsupported();
      try {
        const metadata = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(engine.subarray(offset, offset + length)));
        if (typeof metadata?.collaborative !== 'boolean') return unsupported();
        return metadata.collaborative;
      } catch { return unsupported(); }
    }
  }
  return unsupported();
}

function checksum(bytes: Uint8Array, hash = 2166136261): number {
  for (const byte of bytes) hash = Math.imul(hash ^ byte, 16777619);
  return hash >>> 0;
}

export function encodeWorkbookRecovery(engine: Uint8Array, state: WorkbookCommandState,
  revision: number, versions: [number, string][], refusedSeqs: number[] = [],
  sessionState = { version: revision, dirty: revision > 0 },
  replaySeed: WorkbookRecoveryReplaySeed = noReplaySeed,
  ledger: { refusedRanges?: [number, number][]; refusedOutcomes?: [number, CommandOutcome][];
    historySelections?: WorkbookCommandState['historySelections'] } = {}): Uint8Array {
  replaySeed.capture(state);
  const refusedRanges: [number, number][] = [];
  for (const seq of [...refusedSeqs].sort((left, right) => left - right)) {
    const last = refusedRanges[refusedRanges.length - 1];
    if (last && seq <= last[1] + 1) last[1] = Math.max(last[1], seq);
    else refusedRanges.push([seq, seq]);
  }
  const metadata: RecoveryMetadata = { session: state.identity.session, epoch: state.epoch,
    throughSeq: state.throughSeq, revision, engineVersion: versions.find(([version]) => version === revision)![1],
    versions, bases: [...state.bases], refusedSeqs, refusedRanges: ledger.refusedRanges ?? refusedRanges,
    refusedOutcomes: ledger.refusedOutcomes ?? [], statuses: [...state.statuses], selection: state.selection, readOnly: state.readOnly,
    sheetIds: state.sheetIds, nextSheetId: state.nextSheetId, sheetRevision: state.sheetRevision,
    calculation: state.calculation, formats: [...state.formats], hostVersion: sessionState.version, dirty: sessionState.dirty,
    ...(ledger.historySelections ? { historySelections: ledger.historySelections } : {}) };
  const payload = new TextEncoder().encode(JSON.stringify(metadata));
  if (payload.length > 0xffff_ffff || engine.length === 0 || engine.length > 0xffff_ffff) {
    throw new RangeError('Workbook recovery checkpoint size is out of range');
  }
  const bytes = new Uint8Array(HEADER_BYTES + payload.length + engine.length);
  const header = new DataView(bytes.buffer);
  bytes.set(MAGIC);
  header.setUint8(8, CODEC_VERSION);
  header.setUint32(9, payload.length, true);
  header.setUint32(13, engine.length, true);
  bytes.set(payload, HEADER_BYTES);
  bytes.set(engine, HEADER_BYTES + payload.length);
  header.setUint32(17, checksum(bytes.subarray(HEADER_BYTES), checksum(bytes.subarray(0, 17))), true);
  return bytes;
}

export function decodeWorkbookRecovery(bytes: Uint8Array): WorkbookRecovery {
  const invalid = () => { throw new TypeError('Invalid workbook recovery checkpoint'); };
  if (bytes.length < HEADER_BYTES || !MAGIC.every((byte, index) => bytes[index] === byte)) invalid();
  const header = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const metadataLength = header.getUint32(9, true);
  const engineLength = header.getUint32(13, true);
  if (header.getUint8(8) !== CODEC_VERSION || metadataLength === 0 || engineLength === 0 ||
    HEADER_BYTES + metadataLength + engineLength !== bytes.length ||
    checksum(bytes.subarray(HEADER_BYTES), checksum(bytes.subarray(0, 17))) !== header.getUint32(17, true)) invalid();
  const value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(HEADER_BYTES, HEADER_BYTES + metadataLength)));
  const engine = bytes.subarray(HEADER_BYTES + metadataLength);
  const integer = (number: unknown): number is number => typeof number === 'number' && Number.isSafeInteger(number) && number >= 0;
  const pairs = (entries: unknown, first: (value: unknown) => boolean, second: (value: unknown) => boolean): boolean =>
    Array.isArray(entries) && entries.length <= 256 && entries.every((entry) => Array.isArray(entry) &&
      entry.length === 2 && first(entry[0]) && second(entry[1])) && new Set(entries.map((entry) => entry[0])).size === entries.length;
  const string = (text: unknown) => typeof text === 'string' && text.length > 0;
  const historySelection = (entry: unknown): boolean => {
    if (entry === null) return true;
    if (!entry || typeof entry !== 'object') return false;
    const anchor = entry as NonNullable<WorkbookCommandState['historySelections'][number]>;
    const selection = (value: WorkbookCommandSelection, sheets: string[]) => value && value.range &&
      Array.isArray(sheets) && sheets.length > 0 && sheets.every(string) &&
      new Set(sheets).size === sheets.length && sheets.includes(value.sheetId);
    return Boolean(selection(anchor.before, anchor.beforeSheets) && selection(anchor.after, anchor.afterSheets));
  };
  const refusal = (entry: unknown): boolean => {
    if (!entry || typeof entry !== 'object') return false;
    const outcome = entry as Partial<CommandOutcome>;
    return integer(outcome.seq) && outcome.seq > 0 &&
      (outcome.status === 'refused' || outcome.status === 'cancelled') && string(outcome.reason) &&
      outcome.committed === false && (outcome.affectedSeqs === undefined ||
        (Array.isArray(outcome.affectedSeqs) && outcome.affectedSeqs.every((seq) => integer(seq) && seq > 0)));
  };
  const status = (entry: unknown): boolean => {
    if (!entry || typeof entry !== 'object') return false;
    const value = entry as Partial<CommandStatus>;
    return integer(value.seq) && value.seq > 0 &&
      ['applied', 'noop', 'refused', 'cancelled'].includes(value.status ?? '') &&
      typeof value.committed === 'boolean' &&
      (value.affectedSeqs === undefined || (Array.isArray(value.affectedSeqs) && value.affectedSeqs.every(integer))) &&
      (value.reason === undefined || string(value.reason));
  };
  const format = (value: unknown): boolean => {
    if (!value || typeof value !== 'object') return false;
    const captured = value as Partial<CapturedFormat>;
    return integer(captured.rows) && captured.rows > 0 && integer(captured.columns) && captured.columns > 0 &&
      captured.rows * captured.columns <= 100_000 && Array.isArray(captured.formats) &&
      captured.formats.length === captured.rows * captured.columns;
  };
  if (!value || !string(value.session) || !string(value.engineVersion) || !integer(value.epoch) || value.epoch === 0 ||
    ![value.throughSeq, value.revision, value.hostVersion, value.sheetRevision, value.nextSheetId].every(integer) ||
    typeof value.readOnly !== 'boolean' || typeof value.dirty !== 'boolean' ||
    !Array.isArray(value.sheetIds) || value.sheetIds.length === 0 || !value.sheetIds.every(string) ||
    new Set(value.sheetIds).size !== value.sheetIds.length ||
    value.sheetIds.some((id: string) => id.startsWith('command-sheet:') &&
      (!integer(Number(id.slice(14))) || Number(id.slice(14)) >= value.nextSheetId)) ||
    !pairs(value.versions, integer, string) || !pairs(value.bases, string, integer) ||
    !value.versions.some(([revision, version]: [number, string]) => revision === value.revision && version === value.engineVersion) ||
    !value.bases.some(([version]: [string, number]) => version === value.engineVersion) ||
    !Array.isArray(value.refusedSeqs) || value.refusedSeqs.length > 256 ||
    new Set(value.refusedSeqs).size !== value.refusedSeqs.length ||
    !value.refusedSeqs.every((seq: unknown) => integer(seq) && seq > 0 && seq <= value.throughSeq) ||
    !Array.isArray(value.refusedRanges) || value.refusedRanges.length > 4096 ||
    !value.refusedRanges.every((range: unknown, index: number, ranges: unknown[]) => {
      if (!Array.isArray(range) || range.length !== 2 || !integer(range[0]) || !integer(range[1]) ||
        range[0] === 0 || range[0] > range[1] || range[1] > value.throughSeq) return false;
      const previous = ranges[index - 1];
      return index === 0 || (Array.isArray(previous) && previous[1] < range[0] - 1);
    }) ||
    !value.refusedSeqs.every((seq: number) => value.refusedRanges.some(([start, end]: [number, number]) => seq >= start && seq <= end)) ||
    !pairs(value.statuses, integer, status) ||
    !value.statuses.every(([seq, status]: [number, CommandStatus]) => seq === status.seq && seq <= value.throughSeq &&
      value.refusedSeqs.includes(seq) === (status.status === 'refused' || status.status === 'cancelled')) ||
    !value.statuses.every(([seq, status]: [number, CommandStatus]) =>
      value.refusedRanges.some(([start, end]: [number, number]) => seq >= start && seq <= end) ===
        (status.status === 'refused' || status.status === 'cancelled')) ||
    !pairs(value.refusedOutcomes, integer, refusal) ||
    !value.refusedOutcomes.every(([seq, outcome]: [number, CommandOutcome]) => seq === outcome.seq &&
      value.statuses.some(([retained, status]: [number, CommandStatus]) => retained === seq &&
        status.status === outcome.status && status.reason === outcome.reason) &&
      (outcome.affectedSeqs ?? []).every((affected) =>
        value.statuses.some(([retained, status]: [number, CommandStatus]) => retained === affected &&
          status.status === outcome.status && status.reason === outcome.reason && status.committed === false))) ||
    !pairs(value.formats, string, format) || value.formats.length > 64 ||
    value.historySelections !== undefined && (!Array.isArray(value.historySelections) || !value.historySelections.every(historySelection)) ||
    !value.selection || !value.sheetIds.includes(value.selection.sheetId) ||
    engine.length === 0) invalid();
  if (value.calculation !== null && (!value.calculation || !Number.isFinite(value.calculation.nowSerial) ||
    !integer(value.calculation.randSeed) || value.calculation.randSeed > 0xffff_ffff)) invalid();
  return { ...value, engine };
}

export function restoreWorkbookCommandState(handle: WorkbookHandle, recovery: WorkbookRecovery,
  identity: CommandIdentity, replaySeed: WorkbookRecoveryReplaySeed = noReplaySeed): WorkbookCommandState {
  if (handle.version() !== recovery.engineVersion || handle.sheetCount() !== recovery.sheetIds.length ||
    identity.session !== recovery.session) throw new TypeError('Workbook recovery boundary mismatch');
  const state = createWorkbookCommandState(identity, { epoch: recovery.epoch, version: recovery.engineVersion },
    recovery.selection, recovery.sheetIds, recovery.calculation);
  const target = resolveCommandTarget(handle, recovery.selection, recovery.sheetIds);
  const active = recovery.selection.activeCell;
  if (active && (!Number.isSafeInteger(active.row) || !Number.isSafeInteger(active.col) ||
    active.row < target.start.row || active.row > target.end.row ||
    active.col < target.start.col || active.col > target.end.col)) throw new TypeError('Invalid recovery selection');
  Object.assign(state, { throughSeq: recovery.throughSeq, readOnly: recovery.readOnly,
    nextSheetId: recovery.nextSheetId, sheetRevision: recovery.sheetRevision,
    bases: new Map(recovery.bases), formats: new Map(recovery.formats), historyReset: recovery.historySelections === undefined,
    historySelections: structuredClone(recovery.historySelections ?? []).map((entry) => entry ?? undefined) });
  state.statuses = new Map(recovery.statuses);
  replaySeed.restore(handle);
  projectEditState(handle, state);
  return state;
}
