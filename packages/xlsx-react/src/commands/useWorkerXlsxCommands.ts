import { useLayoutEffect, useRef } from 'react';
import { normalizeRange, StaleProposalError } from '@betteroffice/xlsx';
import type {
  EditResult,
  MergedRange,
  SelectionFormatting,
  WorkbookHandle,
  WorkbookEditPeer,
} from '@betteroffice/xlsx';
import type { MergeAction } from '../components/Toolbar';
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
import { inputRefusal, WorkerInputRefusal, type WorkerInputCoordinator } from './workerInputCoordinator';
import type { XlsxEditorBridge } from './useXlsxCommands';
import type {
  XlsxCommandArgs,
  XlsxCommandFailureCode,
  XlsxCommandId,
  XlsxCommandResult,
} from './types';

export interface WorkerXlsxEditorBridge extends Omit<XlsxEditorBridge, 'handle' | 'coordinator' | 'exportPng'> {
  peer(): WorkbookHandle | null;
  editPeer(): WorkbookEditPeer | null;
  coordinator: WorkerInputCoordinator;
  preview(): Promise<void>;
  recovering?(): boolean;
  refuse?(reason: { message: string }): void;
}

interface EngineRead {
  key: string;
  formatting: SelectionFormatting | null;
  merged: readonly MergedRange[];
  canUndo: boolean;
  canRedo: boolean;
}

const IMMEDIATE: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>([
  'zoom',
  'searchMenus',
  'proposalsPanel',
]);

function rangeA1(peer: WorkbookHandle, range: XlsxSelectionEnvironment): string {
  const from = peer.cell(range.sheet, range.top, range.left).a1;
  const to = peer.cell(range.sheet, range.bottom, range.right).a1;
  return `${from}:${to}`;
}

function readEngine(
  peer: WorkbookHandle,
  sheet: number,
  range: ReturnType<typeof normalizeRange>,
  key: string
): EngineRead {
  const read: EngineRead = { key, formatting: null, merged: [], canUndo: false, canRedo: false };
  try {
    const a1 = rangeA1(peer, { sheet, ...range, merged: [], formatting: null });
    read.formatting = peer.selectionFormatting(sheet, a1);
    read.merged = peer.mergedRanges(sheet, a1);
  } catch {}
  try {
    const history = peer.historyState();
    read.canUndo = history.canUndo;
    read.canRedo = history.canRedo;
  } catch {}
  return read;
}

function mergeOps(
  sheet: number,
  selection: XlsxSelectionEnvironment,
  action: MergeAction
): unknown[] {
  const range = (top: number, left: number, bottom: number, right: number) => ({
    start: { row: top, col: left },
    end: { row: bottom, col: right },
  });
  const ops: unknown[] = [];
  if (action === 'all') {
    ops.push({
      type: 'mergeCells',
      sheet,
      range: range(selection.top, selection.left, selection.bottom, selection.right),
    });
  } else if (action === 'horizontal') {
    for (let row = selection.top; row <= selection.bottom; row++) {
      ops.push({ type: 'mergeCells', sheet, range: range(row, selection.left, row, selection.right) });
    }
  } else if (action === 'vertical') {
    for (let col = selection.left; col <= selection.right; col++) {
      ops.push({ type: 'mergeCells', sheet, range: range(selection.top, col, selection.bottom, col) });
    }
  } else {
    for (const merged of selection.merged) {
      ops.push({
        type: 'unmergeCells',
        sheet,
        range: range(merged.start.row, merged.start.col, merged.end.row, merged.end.col),
      });
    }
  }
  return ops;
}

export function createWorkerXlsxCommandBinding(
  bridge: () => WorkerXlsxEditorBridge
): XlsxCommandBinding {
  let cache: EngineRead | null = null;
  let markApplied: (() => void) | null = null;

  const engine = (executing: boolean): EngineRead | null => {
    const current = bridge();
    const peer = current.peer();
    const view = current.view();
    if (!peer || current.status() !== 'ready') return null;
    const range = view.selection
      ? normalizeRange(view.selection)
      : { top: 0, left: 0, bottom: 0, right: 0 };
    const key = [
      current.generation(),
      current.mutation(),
      view.sheet,
      range.top,
      range.left,
      range.bottom,
      range.right,
    ].join(':');
    if (!executing && cache?.key === key) return cache;
    cache = readEngine(peer, view.sheet, range, key);
    return cache;
  };

  const environment = (executing: boolean): XlsxCommandEnvironment => {
    const current = bridge();
    const view = current.view();
    const read = engine(executing);
    let selection: XlsxCommandEnvironment['selection'] = null;
    if (view.chartSelected) selection = 'chart';
    else if (view.selection) {
      selection = {
        sheet: view.sheet,
        ...normalizeRange(view.selection),
        merged: read?.merged ?? [],
        formatting: read?.formatting ?? null,
      };
    }
    return {
      status: current.status(),
      readOnly: current.readOnly(),
      collaborative: current.collaborative(),
      selection,
      canUndo: read?.canUndo ?? false,
      canRedo: read?.canRedo ?? false,
      pendingInput: current.coordinator.draft !== null || current.coordinator.committed.length > 0,
      zoom: view.zoom,
      paintFormat: view.capturedFormat !== null,
      borderStyle: view.borderStyle ?? read?.formatting?.borderStyle ?? null,
      borderColor: view.borderColor ?? read?.formatting?.borderColor ?? null,
      pngExport: false,
      proposals: view.proposalsAvailable
        ? view.proposals.map((proposal) => ({ id: proposal.id, label: proposal.agentId }))
        : null,
      proposalsPanelOpen: view.proposalsPanelOpen,
      fontFamilies: DEFAULT_FONT_FAMILIES,
      fontSizes: DEFAULT_FONT_SIZES,
      translate: current.translate,
    };
  };

  const perform = <K extends XlsxCommandId>(
    id: K,
    args: XlsxCommandArgs[K],
    env: XlsxCommandEnvironment
  ): XlsxCommandResult | Promise<XlsxCommandResult> => {
    const current = bridge();
    const peer = current.peer();
    const fail = (code: XlsxCommandFailureCode): XlsxCommandResult => ({
      ok: false,
      failure: commandReason(code, env),
    });
    if (id === 'zoom') {
      const scale = (args as XlsxCommandArgs['zoom']).scale;
      if (env.zoom === scale) return { ok: true, status: 'noop' };
      current.setZoom(scale);
      return { ok: true, status: 'executed' };
    }
    const edits = current.editPeer();
    if (!peer || !edits) return fail('editor-unavailable');
    const done = (changed = true): XlsxCommandResult => ({
      ok: true,
      status: changed ? 'executed' : 'noop',
    });
    const attempt = (work: () => XlsxCommandResult): XlsxCommandResult => {
      try {
        return work();
      } catch (error) {
        if (inputRefusal(error)) current.refuse?.({ message: (error as Error).message });
        else current.fail(error);
        return fail('command-failed');
      }
    };
    const edit = (work: () => EditResult) =>
      attempt(() => {
        const result = work();
        markApplied?.();
        current.apply(result);
        return done(result.applied);
      });
    const bound = args as unknown as Record<string, never>;
    const selection = typeof env.selection === 'object' ? env.selection : null;
    const cells = () => {
      if (!selection) throw new Error('no cell selection');
      return { sheet: selection.sheet, range: rangeA1(peer, selection) };
    };
    const style = (patch: Parameters<WorkbookHandle['patchRangeStyle']>[2]) =>
      edit(() => {
        const { sheet, range } = cells();
        return edits.patchRangeStyle(sheet, range, patch);
      });
    const formatting = selection?.formatting ?? null;

    switch (id) {
      case 'bold':
        return style({ bold: formatting?.bold !== true });
      case 'italic':
        return style({ italic: formatting?.italic !== true });
      case 'strikethrough':
        return style({ strikethrough: formatting?.strikethrough !== true });
      case 'paintFormat':
        if (env.paintFormat) {
          current.setCapturedFormat(null, null);
          return done();
        }
        return attempt(() => {
          const { sheet, range } = cells();
          const source = `${sheet}:${selection!.top}:${selection!.left}:${selection!.bottom}:${selection!.right}`;
          current.setCapturedFormat(peer.captureFormat(sheet, range), source);
          return done();
        });
      case 'fontFamily':
        return style({ fontFamily: bound.family });
      case 'fontSize':
        return style({ fontSize: bound.points });
      case 'fontSizeStep': {
        const size = formatting?.fontSize ?? 10;
        const sizes = env.fontSizes;
        const next =
          bound.direction === 'increase'
            ? (sizes.find((entry) => entry > size) ?? size + 1)
            : ([...sizes].reverse().find((entry) => entry < size) ?? Math.max(1, size - 1));
        return style({ fontSize: next });
      }
      case 'textColor':
        return style({ textColor: bound.color });
      case 'fillColor':
        return style({ fillColor: bound.color });
      case 'numberFormat':
        return edit(() => {
          const { sheet, range } = cells();
          return bound.value === 'custom'
            ? edits.setNumberFormat(sheet, range, {
                type: 'custom',
                pattern: formatting?.numberFormatPattern ?? '0.00',
              })
            : edits.setNumberFormat(sheet, range, bound.value);
        });
      case 'decimalPlaces':
        return edit(() => {
          const { sheet, range } = cells();
          return edits.setNumberFormat(
            sheet,
            range,
            bound.direction === 'increase' ? 'increaseDecimal' : 'decreaseDecimal'
          );
        });
      case 'borderPreset':
        return style({
          border: {
            preset: bound.value,
            style: env.borderStyle ?? 'solid',
            color: env.borderColor ?? '#000000',
          },
        });
      case 'borderStyle':
        current.setBorderStyle(bound.value);
        return style({ border: { style: bound.value } });
      case 'borderColor':
        current.setBorderColor(bound.color);
        return style({ border: { color: bound.color } });
      case 'horizontalAlignment':
        return style({ horizontalAlignment: bound.value });
      case 'verticalAlignment':
        return style({ verticalAlignment: bound.value });
      case 'textWrapping':
        return style({ textWrapping: bound.value });
      case 'merge': {
        if (!selection) return fail('cell-selection-required');
        const ops = mergeOps(selection.sheet, selection, bound.value);
        if (ops.length === 0) return done(false);
        return edit(() => edits.applyOps(ops));
      }
      case 'undo':
        return edit(() => edits.undo());
      case 'redo':
        return edit(() => edits.redo());
      case 'save': {
        const generation = current.generation();
        markApplied?.();
        return edits.save().then((bytes) => {
          if (bridge().generation() !== generation) return fail('document-replaced');
          current.deliver(new Uint8Array(bytes));
          return done();
        }).catch((error) => {
          if (bridge().generation() !== generation) return fail('document-replaced');
          current.fail(error);
          return fail('command-failed');
        });
      }
      case 'exportPng':
        return fail('png-unavailable');
      case 'print': {
        const generation = current.generation();
        return current.afterPaint().then((painted) => {
          if (bridge().generation() !== generation) return fail('document-replaced');
          if (!painted) return fail('render-failed');
          window.print();
          return done();
        });
      }
      case 'proposalsPanel': {
        const open = args === null ? !env.proposalsPanelOpen : (bound.open as boolean);
        if (open === env.proposalsPanelOpen) return done(false);
        current.setProposalsPanelOpen(open);
        return done();
      }
      case 'proposalAccept':
        try {
          const result = edits.acceptProposal(bound.proposalId, { force: bound.force });
          markApplied?.();
          current.apply(result);
          current.markStale(bound.proposalId, null);
          current.refreshProposals();
          return done();
        } catch (error) {
          if (error instanceof StaleProposalError) {
            markApplied?.();
            current.markStale(bound.proposalId, error.cells);
            current.refreshProposals();
            return fail('proposal-stale');
          }
          current.fail(error);
          return fail('command-failed');
        }
      case 'proposalReject':
        return attempt(() => {
          const removed = edits.rejectProposal(bound.proposalId);
          markApplied?.();
          current.markStale(bound.proposalId, null);
          current.refreshProposals();
          return removed ? done() : fail('proposal-not-found');
        });
      default:
        return fail('unsupported-command');
    }
  };

  const targetOf = (id: XlsxCommandId): string => {
    if (!isCellCommand(id)) return 'document';
    const view = bridge().view();
    if (view.chartSelected) return 'chart';
    return view.selection ? `${view.sheet}:${JSON.stringify(normalizeRange(view.selection))}` : 'none';
  };

  const binding: XlsxCommandBinding = {
    environment,
    ordered: (id) => {
      const current = bridge();
      if (id === 'save') void current.coordinator.requestHydration('save').catch(() => {});
      return (
        !IMMEDIATE.has(id) ||
        current.coordinator.rejected.some((entry) => entry.generation === current.generation())
      );
    },
    refuse: (reason) => bridge().refuse?.(reason),
    admit: (operation, id) => {
      const readOnly = bridge().readOnly();
      return bridge().coordinator.runAfterPendingInput(async (_, applied) => {
        const current = bridge();
        const generation = current.generation();
        await current.preview();
        applied.check();
        if (bridge().generation() !== generation) throw new XlsxCommandAdmissionError('document-replaced');
        if (current.coordinator.error) throw current.coordinator.error;
        markApplied = applied;
        try {
          const result = await operation(readOnly);
          if (result !== null && typeof result === 'object' && 'ok' in result && result.ok === false) {
            const reason = 'failure' in result ? result.failure as { message: string } : { message: 'Command was refused' };
            applied.refuse(new WorkerInputRefusal(reason.message), id !== undefined && XLSX_COMMAND_DESCRIPTORS[id].mutatesDocument);
          }
          return result;
        } finally { markApplied = null; }
      }, { recover: id === undefined || XLSX_COMMAND_DESCRIPTORS[id].mutatesDocument, barrier: id === 'save' });
    },
    perform,
    capture(id) {
      const current = bridge();
      if (current.status() === 'empty') return null;
      return { generation: current.generation(), target: targetOf(id) };
    },
    resume(origin, id) {
      const current = bridge();
      if (current.generation() !== origin.generation) {
        return 'document-replaced';
      }
      return targetOf(id) === origin.target ? null : 'target-changed';
    },
    chrome: () => ({ i18n: bridge().i18n() }),
    focusEditor: () => bridge().focusGrid(),
  };
  return binding;
}

export function useWorkerXlsxCommands(
  controller: XlsxCommandController, bridge: WorkerXlsxEditorBridge
): void {
  const latest = useRef(bridge);
  latest.current = bridge;
  useLayoutEffect(() => {
    const binding = createWorkerXlsxCommandBinding(() => latest.current);
    controller.attach(binding);
    return () => controller.detach(binding);
  }, [controller]);
  useLayoutEffect(() => controller.refresh());
}
