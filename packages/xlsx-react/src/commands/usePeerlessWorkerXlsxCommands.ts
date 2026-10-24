import { useLayoutEffect, useRef } from 'react';
import { normalizeRange } from '@betteroffice/xlsx';
import type { WorkbookHandle, WorkbookSession, XlsxRangeTarget } from '@betteroffice/xlsx';
import type {
  WorkbookAck, WorkbookCommandSelection, WorkbookCommandFrameRequest, WorkbookEditState,
} from '@betteroffice/xlsx/worker';
import type { WorkbookFormatTokenResult } from '../../../xlsx/src/session/commandProtocol';

import {
  XlsxCommandAdmissionError, type XlsxCommandBinding, type XlsxCommandController,
} from './createXlsxCommandStore';
import {
  commandReason,
  DEFAULT_FONT_FAMILIES,
  DEFAULT_FONT_SIZES,
  isCellCommand,
  type XlsxCommandEnvironment,
  type XlsxSelectionEnvironment,
} from './evaluate';
import { XLSX_COMMAND_DESCRIPTORS } from './descriptors';

import {
  PeerlessWorkerCommandRefusal, type PeerlessWorkerCommandInput, type PeerlessWorkerCommandSession,
  type PeerlessWorkerInputCoordinator,
} from './peerlessWorkerInputCoordinator';

import type {
  XlsxCommandArgs,
  XlsxCommandFailureCode,
  XlsxCommandId,
  XlsxCommandResult,
} from './types';

import type { WorkerXlsxEditorBridge } from './useWorkerXlsxCommands';

const IMMEDIATE: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>([
  'zoom',
  'searchMenus',
  'proposalsPanel',
]);

export interface PeerlessWorkerXlsxEditorBridge extends Omit<WorkerXlsxEditorBridge,
  'peer' | 'editPeer' | 'coordinator' | 'preview' | 'apply' | 'setCapturedFormat' | 'refuse'> {
  session(): (PeerlessWorkerCommandSession & Pick<
    WorkbookSession, 'editState' | 'selection'
  > & Partial<Pick<WorkbookSession, 'call' | 'pngExport'>>) | null;
  deliverPng?(bytes: Uint8Array): void;
  coordinator: PeerlessWorkerInputCoordinator;
  frame(): WorkbookCommandFrameRequest | undefined;
  selectionRevision(): number;
  formatToken(): string | null;
  formatRevision(): number;
  setFormatToken(token: string | null): void;
  selection?(): WorkbookCommandSelection | null;
  refuse?: XlsxCommandBinding['refuse'];
}

export interface PeerlessWorkerXlsxCommandBinding extends XlsxCommandBinding {
  subscribe(listener: () => void): () => void;
  refreshSubscription(): void;
  dispose(): void;
}

function peerlessRange(selection: WorkbookCommandSelection) {
  if (selection.range.kind === 'rowCol') {
    return { top: selection.range.start.row, left: selection.range.start.col,
      bottom: selection.range.end.row, right: selection.range.end.col };
  }
  const match = /^\$?([A-Za-z]+)\$?([1-9]\d*)(?::\$?([A-Za-z]+)\$?([1-9]\d*))?$/.exec(selection.range.a1);
  if (!match) return null;
  const column = (letters: string) => [...letters.toUpperCase()]
    .reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
  return { top: Number(match[2]) - 1, left: column(match[1]),
    bottom: Number(match[4] ?? match[2]) - 1, right: column(match[3] ?? match[1]) };
}

function peerlessA1(selection: XlsxSelectionEnvironment): string {
  const address = (row: number, col: number) => {
    let letters = '';
    for (let value = col + 1; value > 0; value = Math.floor((value - 1) / 26)) {
      letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
    }
    return `${letters}${row + 1}`;
  };
  return `${address(selection.top, selection.left)}:${address(selection.bottom, selection.right)}`;
}

export function createPeerlessWorkerXlsxCommandBinding(
  bridge: () => PeerlessWorkerXlsxEditorBridge
): PeerlessWorkerXlsxCommandBinding {
  const listeners = new Set<() => void>();
  let watched: ReturnType<PeerlessWorkerXlsxEditorBridge['session']> = null;
  let watchedInput: PeerlessWorkerInputCoordinator | null = null;
  let identity = '';
  let pushed = '';
  let formatThroughSeq = 0;
  let offAck = () => {};
  let offInput = () => {};
  let admission: { env: XlsxCommandEnvironment; selection?: WorkbookCommandSelection;
    frame?: WorkbookCommandFrameRequest; selectionRevision?: number } | null = null;

  const selectionOf = (): WorkbookCommandSelection | undefined => {
    const current = bridge();
    const session = current.session();
    if (current.selection) return structuredClone(current.selection() ?? undefined);
    if (session?.selection) return structuredClone(session.selection);
    const view = current.view();
    const sheetId = session?.editState.sheets?.find((sheet) => sheet.index === view.sheet)?.id;
    if (!sheetId || !view.selection) return undefined;
    const range = normalizeRange(view.selection);
    return { sheetId, activeCell: { ...view.selection.anchor }, range: { kind: 'rowCol',
      start: { row: range.top, col: range.left }, end: { row: range.bottom, col: range.right } } };
  };

  const sessionKey = () => {
    const session = bridge().session();
    return `${bridge().generation()}:${session?.commandIdentity.session}:${session?.commandIdentity.incarnation}:${session?.confirmedBasis.epoch}`;
  };

  const projection = () => {
    const session = bridge().session();
    const state = session?.editState;
    return JSON.stringify({
      formatting: state?.formatting, formattingAvailable: state?.formattingAvailable,
      selectionEditable: state?.selectionEditable, mergedRanges: state?.mergedRanges,
      canUndo: state?.canUndo, canRedo: state?.canRedo, readOnly: state?.readOnly,
      activeSheet: state?.activeSheet, sheets: state?.sheets, selection: session?.selection,
    });
  };
  const refreshSubscription = () => {
    const session = bridge().session();
    const coordinator = bridge().coordinator;
    const key = sessionKey();
    if (session === watched && coordinator === watchedInput && identity === key) return;
    offAck(); offInput();
    watched = session;
    watchedInput = coordinator;
    identity = key;
    formatThroughSeq = 0;
    pushed = projection();
    let pending = coordinator.pending;
    offInput = coordinator.subscribe(() => {
      if (bridge().coordinator !== coordinator || bridge().session() !== session || identity !== key || sessionKey() !== key) return;
      const next = coordinator.pending;
      if (next === pending) return;
      pending = next;
      for (const listener of listeners) listener();
    });
    offAck = session?.onAck((ack: WorkbookAck) => {
      if (bridge().session() !== session || identity !== key || sessionKey() !== key) return;
      if (!ack.selection && (!ack.editState || Object.keys(ack.editState).every((key) => key === 'calculation'))) return;
      const next = projection();
      if (next === pushed) return;
      pushed = next;
      for (const listener of listeners) listener();
    }) ?? (() => {});
  };
  const environment = (executing: boolean): XlsxCommandEnvironment => {
    refreshSubscription();
    if (executing && admission) return admission.env;
    const current = bridge();
    const view = current.view();
    const session = current.session();
    const state: Readonly<Partial<WorkbookEditState>> = session?.editState ?? {};
    const restored = current.selection ? current.selection() : session?.selection;
    const range = restored ? peerlessRange(restored) :
      current.selection ? null : view.selection ? normalizeRange(view.selection) : null;
    const sheet = restored ? state.sheets?.find((entry) => entry.id === restored.sheetId)?.index ?? view.sheet : view.sheet;
    return {
      status: current.status(), readOnly: current.readOnly() || current.coordinator.readOnly || state.readOnly === true,
      collaborative: current.collaborative(),
      selection: view.chartSelected ? 'chart' : range ? {
        sheet, ...range, merged: state.mergedRanges ?? [],
        formatting: state.formattingAvailable === false ? null : state.formatting ?? null,
      } : null,
      canUndo: current.coordinator.pending || (state.canUndo ?? false),
      canRedo: current.coordinator.pending || (state.canRedo ?? false),
      pendingInput: current.coordinator.draft !== null || current.coordinator.committed.length > 0,
      zoom: view.zoom, paintFormat: current.formatToken() !== null,
      borderStyle: view.borderStyle ?? state.formatting?.borderStyle ?? null,
      borderColor: view.borderColor ?? state.formatting?.borderColor ?? null,
      pngExport: session?.pngExport === true,
      proposals: view.proposalsAvailable ? view.proposals.map((proposal) => ({ id: proposal.id, label: proposal.agentId })) : null,
      proposalsPanelOpen: view.proposalsPanelOpen, fontFamilies: DEFAULT_FONT_FAMILIES,
      fontSizes: DEFAULT_FONT_SIZES, translate: current.translate,
      hostDisabled: state.selectionEditable === false ? new Set(
        (Object.keys(XLSX_COMMAND_DESCRIPTORS) as XlsxCommandId[]).filter(isCellCommand)
      ) : undefined,
    };
  };
  const targetOf = (id: XlsxCommandId, env = environment(false)) => {
    if (!isCellCommand(id)) return 'document';
    const selection = env.selection;
    if (selection === 'chart') return 'chart';
    return selection ? `${selection.sheet}:${selection.top}:${selection.left}:${selection.bottom}:${selection.right}` : 'none';
  };
  const perform = async <K extends XlsxCommandId>(
    id: K, args: XlsxCommandArgs[K], env: XlsxCommandEnvironment
  ): Promise<XlsxCommandResult> => {
    const current = bridge();
    const generation = current.generation();
    const session = current.session();
    const admitted = admission;
    const selection = admitted ? admitted.selection : selectionOf();
    const done = (changed = true): XlsxCommandResult => ({ ok: true, status: changed ? 'executed' : 'noop' });
    const fail = (code: XlsxCommandFailureCode): XlsxCommandResult => ({ ok: false, failure: commandReason(code, env) });
    if (id === 'zoom') {
      const scale = (args as XlsxCommandArgs['zoom']).scale;
      if (scale === env.zoom) return done(false);
      current.setZoom(scale);
      return done();
    }
    if (id === 'proposalsPanel') {
      const open = args === null ? !env.proposalsPanelOpen : (args as XlsxCommandArgs['proposalsPanel'])!.open;
      if (open === env.proposalsPanelOpen) return done(false);
      current.setProposalsPanelOpen(open);
      return done();
    }
    if (!session) return fail('editor-unavailable');
    if (id === 'paintFormat') {
      const armed = current.formatToken() !== null;
      current.setFormatToken(null);
      if (armed) return done();
    }
    const formatRevision = current.formatRevision();
    const cells = typeof env.selection === 'object' ? env.selection : null;
    const sheetId = cells ? session.editState.sheets?.find((sheet) => sheet.index === cells.sheet)?.id : undefined;
    const target: XlsxRangeTarget | undefined = cells && sheetId ? {
      sheetId, range: { kind: 'rowCol', start: { row: cells.top, col: cells.left },
        end: { row: cells.bottom, col: cells.right } },
    } : undefined;
    const argsRecord = args as unknown as Record<string, never>;
    const style = (patch: Parameters<WorkbookHandle['patchRangeStyle']>[2]): PeerlessWorkerCommandInput => {
      if (!cells || !target) throw new XlsxCommandAdmissionError('editor-unavailable');
      return { kind: 'patchRangeStyle', payload: {
        args: [cells.sheet, peerlessA1(cells), patch], target, selection,
      } };
    };
    const worker = async <T>(operation: () => Promise<T>): Promise<T> => {
      try { return await operation(); } catch (error) {
        if (bridge().generation() === generation && bridge().session() === session &&
          !(error instanceof PeerlessWorkerCommandRefusal) && !(error instanceof XlsxCommandAdmissionError) &&
          !current.coordinator.error) current.coordinator.fail(error);
        throw error;
      }
    };
    try {
      if (id === 'save') {
        const bytes = await worker(() => current.coordinator.save());
        if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
        current.deliver(bytes);
        return done();
      }
      if (id === 'print') {
        await worker(() => current.coordinator.whenPresented());
        if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
        window.print();
        return done();
      }
      if (id === 'exportPng') {
        const frame = admitted ? admitted.frame : current.frame();
        if (!session.pngExport || !session.call) return fail('png-unavailable');
        if (!frame || !current.deliverPng) return fail('editor-unavailable');
        const png = await session.call.renderPng(frame.viewport, { sheet: frame.sheet });
        if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
        current.deliverPng(png);
        return done();
      }
      let input: PeerlessWorkerCommandInput;
      switch (id) {
        case 'bold': case 'italic': case 'strikethrough': case 'fontSizeStep':
        case 'merge': case 'paintFormat':
          if (!target) return fail('cell-selection-required');
          input = { kind: 'applyOps', payload: { toolbar: { id, args, target }, selection } };
          break;
        case 'fontFamily': input = style({ fontFamily: argsRecord.family }); break;
        case 'fontSize': input = style({ fontSize: argsRecord.points }); break;
        case 'textColor': input = style({ textColor: argsRecord.color }); break;
        case 'fillColor': input = style({ fillColor: argsRecord.color }); break;
        case 'horizontalAlignment': input = style({ horizontalAlignment: argsRecord.value }); break;
        case 'verticalAlignment': input = style({ verticalAlignment: argsRecord.value }); break;
        case 'textWrapping': input = style({ textWrapping: argsRecord.value }); break;
        case 'borderPreset': input = style({ border: {
          preset: argsRecord.value, style: env.borderStyle ?? 'solid', color: env.borderColor ?? '#000000',
        } }); break;
        case 'borderStyle':
          current.setBorderStyle(argsRecord.value);
          input = style({ border: { style: argsRecord.value } }); break;
        case 'borderColor':
          current.setBorderColor(argsRecord.color);
          input = style({ border: { color: argsRecord.color } }); break;
        case 'numberFormat': case 'decimalPlaces':
          if (!cells || !target) return fail('cell-selection-required');
          if (id === 'numberFormat' && argsRecord.value === 'custom') {
            input = { kind: 'applyOps', payload: { toolbar: { id, args, target }, selection } };
          } else {
            input = { kind: 'setNumberFormat', payload: { args: [cells.sheet, peerlessA1(cells),
              id === 'decimalPlaces' ? argsRecord.direction === 'increase' ? 'increaseDecimal' : 'decreaseDecimal' : argsRecord.value,
            ], target, selection } };
          }
          break;
        case 'undo': case 'redo': input = { kind: id, payload: { args: [] } }; break;
        case 'proposalAccept': input = { kind: 'acceptProposal', payload: {
          args: [argsRecord.proposalId, { force: argsRecord.force }],
        } }; break;
        case 'proposalReject': input = { kind: 'rejectProposal', payload: { args: [argsRecord.proposalId] } }; break;
        default: return fail('unsupported-command');
      }
      const ticket = current.coordinator.command(input, { frame: admitted?.frame,
        selectionRevision: admitted?.selectionRevision });
      const status = await worker(() => ticket.result);
      if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
      if (id === 'paintFormat' && ticket.seq !== null && ticket.seq >= formatThroughSeq &&
        bridge().formatRevision() === formatRevision) {
        const result = status.result as WorkbookFormatTokenResult | undefined;
        if (typeof result?.formatToken !== 'string') {
          const error = new Error('Worker format capture omitted its token');
          current.coordinator.fail(error);
          throw error;
        }
        current.setFormatToken(result.formatToken);
        formatThroughSeq = ticket.seq;
      }
      if (id === 'proposalAccept' || id === 'proposalReject') {
        current.markStale(argsRecord.proposalId, null);
        current.refreshProposals();
      }
      if (id === 'proposalReject' && status.result === false) return fail('proposal-not-found');
      return done(status.status !== 'noop');
    } catch (error) {
      if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
      if (error instanceof PeerlessWorkerCommandRefusal) {
        if (id === 'proposalAccept' && error.status.reason === 'stale-proposal') {
          const result = error.status.result as { cells?: string[] } | undefined;
          current.markStale(argsRecord.proposalId, result?.cells ?? []);
          current.refreshProposals();
          return fail('proposal-stale');
        }
        if (error.status.reason === 'read-only') return fail('read-only');
        if (id === 'proposalReject' && error.status.reason === 'missing-target') return fail('proposal-not-found');
        return fail('input-failed');
      }
      if (error instanceof XlsxCommandAdmissionError) return fail(error.code);
      if (error instanceof Error && error.name === 'WorkbookPngUnavailableError') return fail('png-unavailable');
      return fail('command-failed');
    }
  };
  return {
    requestSave: (save) => bridge().requestSave?.(save) ?? save(),
    environment, ordered: (id) => !IMMEDIATE.has(id), perform,
    admit(operation, id) {
      const env = environment(false);
      const captured = { env, selection: selectionOf(), frame: structuredClone(bridge().frame()),
        selectionRevision: bridge().selectionRevision() };
      return bridge().coordinator.runAfterPendingInput(() => {
        const previous = admission;
        admission = captured;
        try { return operation(env.readOnly); }
        finally { admission = previous; }
      }, id === 'exportPng');
    },
    capture(id) {
      if (bridge().status() === 'empty') return null;
      return { generation: bridge().generation(), target: targetOf(id) };
    },
    resume(origin, id) {
      if (bridge().generation() !== origin.generation) return 'document-replaced';
      return targetOf(id, admission?.env) === origin.target ? null : 'target-changed';
    },
    refuse: (reason) => bridge().refuse?.(reason),
    chrome: () => ({ i18n: bridge().i18n() }), focusEditor: () => bridge().focusGrid(),
    subscribe(listener) {
      listeners.add(listener);
      refreshSubscription();
      return () => { listeners.delete(listener); };
    },
    refreshSubscription,
    dispose() { offAck(); offInput(); listeners.clear(); watched = null; watchedInput = null; identity = ''; },
  };
}

export function usePeerlessWorkerXlsxCommands(
  controller: XlsxCommandController, bridge: PeerlessWorkerXlsxEditorBridge
): void {
  const latest = useRef(bridge);
  latest.current = bridge;
  const binding = useRef<PeerlessWorkerXlsxCommandBinding | null>(null);
  useLayoutEffect(() => {
    const next = createPeerlessWorkerXlsxCommandBinding(() => latest.current);
    binding.current = next;
    controller.attach(next);
    const unsubscribe = next.subscribe(() => controller.refresh());
    return () => {
      unsubscribe(); next.dispose(); controller.detach(next);
      binding.current = null;
    };
  }, [controller]);
  useLayoutEffect(() => { binding.current?.refreshSubscription(); controller.refresh(); });
}
