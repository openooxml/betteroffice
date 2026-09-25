import { useLayoutEffect, useRef } from 'react';
import { normalizeRange, StaleProposalError } from '@betteroffice/xlsx';
import type {
  CapturedFormat,
  EditResult,
  MergedRange,
  Proposal,
  Selection,
  SelectionFormatting,
  WorkbookHandle,
} from '@betteroffice/xlsx';
import type { TFunction, Translations } from '@betteroffice/xlsx-i18n';
import type { BorderStyle, MergeAction } from '../components/Toolbar';
import type { XlsxCommandBinding, XlsxCommandController } from './createXlsxCommandStore';
import {
  commandReason,
  DEFAULT_FONT_FAMILIES,
  DEFAULT_FONT_SIZES,
  isCellCommand,
  type XlsxCommandEnvironment,
  type XlsxSelectionEnvironment,
} from './evaluate';
import type { InputCoordinator } from './inputCoordinator';
import type {
  XlsxCommandArgs,
  XlsxCommandFailureCode,
  XlsxCommandId,
  XlsxCommandResult,
} from './types';

/** What the command binding reads from the editor at the moment it is asked. */
export interface XlsxEditorView {
  sheet: number;
  selection: Selection | null;
  chartSelected: boolean;
  zoom: number;
  capturedFormat: CapturedFormat | null;
  borderStyle: BorderStyle | undefined;
  borderColor: string | undefined;
  proposals: readonly Proposal[];
  proposalsAvailable: boolean;
  proposalsPanelOpen: boolean;
  pngExport: boolean;
}

/** The editor operations commands are built from; every member reads live state. */
export interface XlsxEditorBridge {
  handle(): WorkbookHandle | null;
  status(): XlsxCommandEnvironment['status'];
  readOnly(): boolean;
  collaborative(): boolean;
  /** Bumped by every committed change, so cached engine reads stay current. */
  mutation(): number;
  generation(): number;
  view(): XlsxEditorView;
  coordinator: InputCoordinator;
  i18n(): Translations | undefined;
  translate: TFunction;
  apply(result: EditResult): void;
  fail(error: unknown): void;
  setZoom(scale: number): void;
  setProposalsPanelOpen(open: boolean): void;
  setCapturedFormat(format: CapturedFormat | null, source: string | null): void;
  setBorderStyle(style: BorderStyle): void;
  setBorderColor(color: string): void;
  markStale(proposalId: string, cells: string[] | null): void;
  refreshProposals(): void;
  deliver(bytes: Uint8Array): void;
  exportPng(): void;
  /** True once the canvas paints every committed change; false if it does not in time. */
  afterPaint(): Promise<boolean>;
  focusGrid(): void;
}

interface EngineRead {
  key: string;
  formatting: SelectionFormatting | null;
  merged: readonly MergedRange[];
  canUndo: boolean;
  canRedo: boolean;
}

/** Commands that change only the view, so they need not wait for input. */
const IMMEDIATE: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>([
  'zoom',
  'searchMenus',
  'proposalsPanel',
]);

function rangeA1(handle: WorkbookHandle, range: XlsxSelectionEnvironment): string {
  const from = handle.cell(range.sheet, range.top, range.left).a1;
  const to = handle.cell(range.sheet, range.bottom, range.right).a1;
  return `${from}:${to}`;
}

function readEngine(
  handle: WorkbookHandle,
  sheet: number,
  range: ReturnType<typeof normalizeRange>,
  key: string
): EngineRead {
  const read: EngineRead = { key, formatting: null, merged: [], canUndo: false, canRedo: false };
  try {
    const a1 = rangeA1(handle, { sheet, ...range, merged: [], formatting: null });
    read.formatting = handle.selectionFormatting(sheet, a1);
    read.merged = handle.mergedRanges(sheet, a1);
  } catch {}
  try {
    const history = handle.historyState();
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

/**
 * Binds the editor's command store to `bridge`, which the editor refreshes
 * every render; snapshots re-evaluate after every render.
 */
export function useXlsxCommandBinding(
  controller: XlsxCommandController,
  bridge: XlsxEditorBridge
): void {
  const bridgeRef = useRef(bridge);
  bridgeRef.current = bridge;
  const cache = useRef<EngineRead | null>(null);

  useLayoutEffect(() => {
    const engine = (executing: boolean): EngineRead | null => {
      const current = bridgeRef.current;
      const handle = current.handle();
      const view = current.view();
      if (!handle || current.status() !== 'ready') return null;
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
      if (!executing && cache.current?.key === key) return cache.current;
      cache.current = readEngine(handle, view.sheet, range, key);
      return cache.current;
    };

    const environment = (executing: boolean): XlsxCommandEnvironment => {
      const current = bridgeRef.current;
      const view = current.view();
      const read = engine(executing);
      let selection: XlsxCommandEnvironment['selection'] = null;
      if (view.chartSelected) selection = 'chart';
      else if (view.selection && read) {
        selection = {
          sheet: view.sheet,
          ...normalizeRange(view.selection),
          merged: read.merged,
          formatting: read.formatting,
        };
      }
      return {
        status: current.status(),
        readOnly: current.readOnly(),
        collaborative: current.collaborative(),
        selection,
        canUndo: read?.canUndo ?? false,
        canRedo: read?.canRedo ?? false,
        pendingInput: current.coordinator.draft !== null,
        zoom: view.zoom,
        paintFormat: view.capturedFormat !== null,
        borderStyle: view.borderStyle ?? read?.formatting?.borderStyle ?? null,
        borderColor: view.borderColor ?? read?.formatting?.borderColor ?? null,
        pngExport: view.pngExport,
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
      const current = bridgeRef.current;
      const handle = current.handle();
      const fail = (code: XlsxCommandFailureCode): XlsxCommandResult => ({
        ok: false,
        failure: commandReason(code, env),
      });
      if (!handle) return fail('editor-unavailable');
      const done = (changed = true): XlsxCommandResult => ({
        ok: true,
        status: changed ? 'executed' : 'noop',
      });
      const attempt = (work: () => XlsxCommandResult): XlsxCommandResult => {
        try {
          return work();
        } catch (error) {
          current.fail(error);
          return fail('execution-failed');
        }
      };
      const edit = (work: () => EditResult) =>
        attempt(() => {
          const result = work();
          current.apply(result);
          return done(result.applied);
        });
      const bound = args as unknown as Record<string, never>;
      const selection = typeof env.selection === 'object' ? env.selection : null;
      const cells = () => {
        if (!selection) throw new Error('no cell selection');
        return { sheet: selection.sheet, range: rangeA1(handle, selection) };
      };
      const style = (patch: Parameters<WorkbookHandle['patchRangeStyle']>[2]) =>
        edit(() => {
          const { sheet, range } = cells();
          return handle.patchRangeStyle(sheet, range, patch);
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
            current.setCapturedFormat(handle.captureFormat(sheet, range), source);
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
              ? handle.setNumberFormat(sheet, range, {
                  type: 'custom',
                  pattern: formatting?.numberFormatPattern ?? '0.00',
                })
              : handle.setNumberFormat(sheet, range, bound.value);
          });
        case 'decimalPlaces':
          return edit(() => {
            const { sheet, range } = cells();
            return handle.setNumberFormat(
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
          return edit(() => handle.applyOps(ops));
        }
        case 'undo':
          return edit(() => handle.undo());
        case 'redo':
          return edit(() => handle.redo());
        case 'save':
          return attempt(() => {
            current.deliver(handle.save());
            return done();
          });
        case 'exportPng':
          return attempt(() => {
            current.exportPng();
            return done();
          });
        case 'print': {
          const generation = current.generation();
          return current.afterPaint().then((painted) => {
            if (bridgeRef.current.generation() !== generation) return fail('document-replaced');
            if (!painted) return fail('render-failed');
            window.print();
            return done();
          });
        }
        case 'zoom':
          if (env.zoom === bound.scale) return done(false);
          current.setZoom(bound.scale);
          return done();
        case 'proposalsPanel': {
          const open = args === null ? !env.proposalsPanelOpen : (bound.open as boolean);
          if (open === env.proposalsPanelOpen) return done(false);
          current.setProposalsPanelOpen(open);
          return done();
        }
        case 'proposalAccept':
          try {
            current.apply(handle.acceptProposal(bound.proposalId, { force: bound.force }));
            current.markStale(bound.proposalId, null);
            current.refreshProposals();
            return done();
          } catch (error) {
            if (error instanceof StaleProposalError) {
              current.markStale(bound.proposalId, error.cells);
              current.refreshProposals();
              return fail('proposal-stale');
            }
            current.fail(error);
            return fail('execution-failed');
          }
        case 'proposalReject':
          return attempt(() => {
            const removed = handle.rejectProposal(bound.proposalId);
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
      const view = bridgeRef.current.view();
      if (view.chartSelected) return 'chart';
      return view.selection ? `${view.sheet}:${JSON.stringify(normalizeRange(view.selection))}` : 'none';
    };

    const binding: XlsxCommandBinding = {
      environment,
      ordered: (id) => {
        const current = bridgeRef.current;
        return (
          !IMMEDIATE.has(id) ||
          current.coordinator.rejected.some((entry) => entry.generation === current.generation())
        );
      },
      admit: (operation) => bridgeRef.current.coordinator.runAfterPendingInput(operation),
      perform,
      capture(id) {
        const current = bridgeRef.current;
        if (!current.handle()) return null;
        return { generation: current.generation(), target: targetOf(id) };
      },
      resume(origin, id) {
        const current = bridgeRef.current;
        if (!current.handle() || current.generation() !== origin.generation) {
          return 'document-replaced';
        }
        return targetOf(id) === origin.target ? null : 'target-changed';
      },
      chrome: () => ({ i18n: bridgeRef.current.i18n() }),
      focusEditor: () => bridgeRef.current.focusGrid(),
    };
    controller.attach(binding);
    return () => controller.detach(binding);
  }, [controller]);

  useLayoutEffect(() => {
    controller.refresh();
  });
}
