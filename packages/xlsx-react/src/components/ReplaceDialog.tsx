import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import type {
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
} from 'react';
import type {
  XlsxCellAddress,
  XlsxEditPreview,
  XlsxEditRequest,
  XlsxEditResult,
  XlsxEditStep,
  XlsxFindMatch,
  XlsxFindRequest,
  XlsxFindResult,
  XlsxRangeAddress,
  XlsxReadResult,
  XlsxReplaceStats,
  XlsxSheetEntry,
  XlsxValidationResult,
} from '@betteroffice/xlsx';
import { useTranslation } from '../i18n';
import { resolvedLocale } from '../i18n/LocaleContext';
import { toolbarTheme } from '../../../../shared/react-toolbar/theme';

/** One sheet the dialog may search when the scope is the whole workbook. */
export interface ReplaceDialogSheet {
  sheetId: string;
  name: string;
  /** From the engine's sheet catalog; only the sheet count shown for a workbook search reads it. */
  hidden?: boolean;
  /** As `hidden`; omitted means a worksheet. */
  kind?: XlsxSheetEntry['kind'];
}

/** `sheets` with visibility and kind from a sheet catalog read; unchanged when the read refused. */
export function withSheetCatalog(
  sheets: readonly ReplaceDialogSheet[],
  catalog: XlsxReadResult,
): ReplaceDialogSheet[] {
  if (!catalog.ok) return [...sheets];
  return sheets.map((sheet) => {
    const entry = catalog.sheets.find((candidate) => candidate.sheetId === sheet.sheetId);
    return entry
      ? {
        ...sheet,
        hidden: entry.visibility === 'hidden' || entry.visibility === 'veryHidden',
        kind: entry.kind,
      }
      : sheet;
  });
}

/** The range a replace dialog works on; the caller resolves the sheet and selection. */
export interface ReplaceDialogTarget {
  sheetId: string;
  range: XlsxRangeAddress;
  /** Human-readable range, such as `A1:C20`. */
  label: string;
  /** Scope selected on open; derived from the range when omitted. */
  defaultScope?: 'selection' | 'sheet' | 'workbook';
  /** Every sheet in order; lets the dialog offer this sheet or all sheets. */
  sheets?: readonly ReplaceDialogSheet[];
  /** The grid's active cell on `sheetId`; Find next starts after it. Defaults to the range start. */
  activeCell?: { row: number; col: number };
}

/** The engine calls the dialog needs; both editors provide them over their session. */
export interface ReplaceDialogApi {
  version(): string | Promise<string>;
  find(request: XlsxFindRequest): XlsxFindResult | Promise<XlsxFindResult>;
  validate(request: XlsxEditRequest): XlsxValidationResult | Promise<XlsxValidationResult>;
  apply(request: XlsxEditRequest): XlsxEditResult | Promise<XlsxEditResult>;
  /** Selects a match in the grid without moving focus away from the dialog. */
  reveal?(cell: XlsxCellAddress): void | Promise<void>;
}

type Phase =
  | { kind: 'idle' }
  | { kind: 'busy' }
  | { kind: 'preview'; preview: XlsxEditPreview; stats: XlsxReplaceStats | null }
  | { kind: 'applied'; preview: XlsxEditPreview; stats: XlsxReplaceStats | null }
  | { kind: 'refused'; message: string };

type ReplaceScope = 'selection' | 'sheet' | 'workbook';

interface FindCache {
  key: string;
  /** Sorted by sheet order, then row, then column. */
  matches: XlsxFindMatch[];
  sheetOrder: string[];
  truncated: boolean;
}

interface CellPosition {
  sheetId: string;
  row: number;
  col: number;
}

const DIALOG_WIDTH = 420;
const BUTTON = { height: 30, padding: '0 12px', borderRadius: 4 } as const;
const NAV_BUTTON = { ...BUTTON, width: 30, padding: 0 } as const;
const FIND_LIMIT = 10_000;
/** Replace edits each cell's entered text, as Excel does; Look in only steers Find. */
const REPLACE_LOOK_IN = 'formulas';
const SHEET_LAST_ROW = 1_048_575;
const SHEET_LAST_COLUMN = 16_383;
const WHOLE_SHEET_RANGE: XlsxRangeAddress = {
  kind: 'rowCol',
  start: { row: 0, col: 0 },
  end: { row: SHEET_LAST_ROW, col: SHEET_LAST_COLUMN },
};

function scopeTargets(
  target: ReplaceDialogTarget,
  scope: ReplaceScope,
): { sheetId: string; range: XlsxRangeAddress }[] {
  if (scope === 'selection') {
    return [{ sheetId: target.sheetId, range: target.range }];
  }
  if (scope === 'sheet') {
    return [{ sheetId: target.sheetId, range: WHOLE_SHEET_RANGE }];
  }
  return workbookSheets(target).map((sheet) => ({ sheetId: sheet.sheetId, range: WHOLE_SHEET_RANGE }));
}

/** Every sheet in order; the engine skips the ones a workbook search leaves out. */
function workbookSheets(target: ReplaceDialogTarget): readonly ReplaceDialogSheet[] {
  return target.sheets?.length ? target.sheets : [{ sheetId: target.sheetId, name: target.label }];
}

/** The sheets the engine's catalog says a workbook search covers. */
function searchedSheetCount(target: ReplaceDialogTarget): number {
  return workbookSheets(target)
    .filter((sheet) => !sheet.hidden && (sheet.kind ?? 'worksheet') === 'worksheet').length;
}

function sumReplaceStats(
  previews: readonly XlsxEditPreview[],
  sheets: readonly ReplaceDialogSheet[] = [],
): XlsxReplaceStats | null {
  const total: XlsxReplaceStats = {
    matched: 0, replaced: 0, skippedFormulas: 0, skippedLocked: 0,
  };
  let skippedValues = 0;
  const skippedCells: string[] = [];
  const refusedCells: { cell: string; reason: string }[] = [];
  let refused = 0;
  let found = false;
  for (const preview of previews) {
    if (!preview.replace) continue;
    found = true;
    total.matched += preview.replace.matched;
    total.replaced += preview.replace.replaced;
    total.skippedFormulas += preview.replace.skippedFormulas;
    total.skippedLocked += preview.replace.skippedLocked;
    skippedValues += preview.replace.skippedValues ?? 0;
    const sheetName = previews.length > 1
      ? sheets.find((sheet) => sheet.sheetId === preview.target.sheetId)?.name
      : undefined;
    for (const cell of preview.replace.skippedCells ?? []) {
      skippedCells.push(sheetName ? `${sheetName}!${cell}` : cell);
    }
    refusedCells.push(...(preview.replace.refusedCells ?? []));
    refused += refusedTotal(preview.replace);
  }
  return found ? { ...total, skippedValues, skippedCells, refusedCells, refused } : null;
}

function refusedTotal(stats: XlsxReplaceStats): number {
  return stats.refused ?? stats.refusedCells?.length ?? 0;
}

const SKIPPED_CELLS_SHOWN = 5;

function skippedCellList(stats: XlsxReplaceStats): string {
  const shown = (stats.skippedCells ?? []).slice(0, SKIPPED_CELLS_SHOWN).join(', ');
  return (stats.skippedValues ?? 0) > SKIPPED_CELLS_SHOWN ? `${shown}, …` : shown;
}

/** The engine's reasons for the first refused cells and how many more it refused. */
function refusedCellList(stats: XlsxReplaceStats): string {
  const listed = (stats.refusedCells ?? []).slice(0, SKIPPED_CELLS_SHOWN);
  const shown = listed.map((entry) => entry.reason).join('; ');
  const more = refusedTotal(stats) - listed.length;
  return more > 0 ? `${shown}; … +${more}` : shown;
}

function refusalMessage(result: { failure: { message: string } }): string {
  return result.failure.message;
}

function searchKey(
  target: ReplaceDialogTarget,
  scope: ReplaceScope,
  query: string,
  lookIn: 'values' | 'formulas',
  entireCell: boolean,
  matchCase: boolean,
  version: string,
): string {
  const sheetIds = scopeTargets(target, scope).map((range) => range.sheetId).join(',');
  return [
    target.sheetId, target.label, sheetIds, scope, query, lookIn, String(entireCell), String(matchCase),
    version,
  ].join('\u0000');
}

function activeCellOf(target: ReplaceDialogTarget): CellPosition {
  const cell = target.activeCell
    ?? (target.range.kind === 'rowCol' ? target.range.start : { row: 0, col: 0 });
  return { sheetId: target.sheetId, row: cell.row, col: cell.col };
}

function compareCells(sheetOrder: readonly string[], a: CellPosition, b: CellPosition): number {
  const rank = (sheetId: string) => {
    const index = sheetOrder.indexOf(sheetId);
    return index < 0 ? sheetOrder.length : index;
  };
  return rank(a.sheetId) - rank(b.sheetId) || a.row - b.row || a.col - b.col;
}

/** The first match after `from` (or the last before it), wrapping around the workbook. */
function stepFrom(found: FindCache, from: CellPosition, direction: 1 | -1): number {
  const { matches, sheetOrder } = found;
  if (direction > 0) {
    const index = matches.findIndex((entry) => compareCells(sheetOrder, entry.cell, from) > 0);
    return index < 0 ? 0 : index;
  }
  for (let index = matches.length - 1; index >= 0; index -= 1) {
    if (compareCells(sheetOrder, matches[index].cell, from) < 0) return index;
  }
  return matches.length - 1;
}

/** A single active cell means "search everything"; an explicit multi-cell selection means itself. */
function defaultScopeFor(target: ReplaceDialogTarget | null): ReplaceScope {
  if (!target) return 'workbook';
  if (target.defaultScope) return target.defaultScope;
  const { range } = target;
  if (range.kind === 'rowCol') {
    return range.start.row === range.end.row && range.start.col === range.end.col
      ? 'workbook'
      : 'selection';
  }
  return range.a1.includes(':') ? 'selection' : 'workbook';
}

/**
 * Find and replace over one range. Find next/previous selects one matching cell at a time;
 * Replace rewrites only that cell, or only selects the next match when that cell does not match.
 * Replace all still stages one `replaceText` step.
 */
export function ReplaceDialog({
  open,
  target,
  api,
  readOnly = false,
  onClose,
  onApplied,
}: {
  open: boolean;
  target: ReplaceDialogTarget | null;
  api: ReplaceDialogApi;
  readOnly?: boolean;
  onClose(): void;
  onApplied?(result: XlsxEditResult): void;
}) {
  const { t, lang } = useTranslation();
  const [lookIn, setLookIn] = useState<'values' | 'formulas'>('values');
  const [scope, setScope] = useState<ReplaceScope>('selection');
  const [entireCell, setEntireCell] = useState(false);
  const [matchCase, setMatchCase] = useState(false);
  const [phase, setPhase] = useState<Phase>({ kind: 'idle' });
  const [match, setMatch] = useState<{ index: number; total: number; truncated: boolean } | null>(null);
  const [position, setPosition] = useState({ x: 0, y: 0 });
  const [dragging, setDragging] = useState(false);
  const dragOrigin = useRef<{ x: number; y: number; baseX: number; baseY: number } | null>(null);
  const cache = useRef<FindCache | null>(null);
  const cursor = useRef<CellPosition | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const replaceRef = useRef<HTMLInputElement>(null);
  const titleId = useId();
  const findId = useId();
  const replaceId = useId();
  const lookInId = useId();
  const entireId = useId();
  const matchCaseId = useId();
  const scopeId = useId();

  useLayoutEffect(() => {
    if (!open) return;
    cache.current = null;
    cursor.current = null;
    setMatch(null);
    setPhase({ kind: 'idle' });
    setScope(defaultScopeFor(target));
    setPosition({ x: 0, y: 0 });
    setDragging(false);
    inputRef.current?.focus();
    inputRef.current?.select();
  }, [open]);

  useEffect(() => {
    if (!dragging) return;
    const move = (event: MouseEvent) => {
      const origin = dragOrigin.current;
      if (!origin) return;
      setPosition({
        x: origin.baseX + event.clientX - origin.x,
        y: origin.baseY + event.clientY - origin.y,
      });
    };
    const end = () => {
      dragOrigin.current = null;
      setDragging(false);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', end);
    return () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', end);
    };
  }, [dragging]);

  if (!open) return null;
  const busy = phase.kind === 'busy';
  const ready = target !== null && !busy && !readOnly;
  const stats = phase.kind === 'preview' || phase.kind === 'applied' ? phase.stats : null;
  const activeSheetName = target?.sheets?.find((sheet) => sheet.sheetId === target.sheetId)?.name
    ?? target?.label ?? '';
  const workbookSheetCount = target ? searchedSheetCount(target) : 1;

  const findMatches = async (
    query: string,
    mode: 'values' | 'formulas' = lookIn,
  ): Promise<FindCache | null> => {
    if (!target) return null;
    const version = await Promise.resolve(api.version());
    const key = searchKey(target, scope, query, mode, entireCell, matchCase, version);
    if (cache.current?.key === key) return cache.current;
    const ranges = scopeTargets(target, scope);
    const result = await Promise.resolve(api.find(scope === 'workbook'
      ? {
        text: query,
        lookIn: mode,
        entireCell,
        matchCase,
        limit: FIND_LIMIT,
      }
      : {
        text: query,
        range: ranges[0],
        lookIn: mode,
        entireCell,
        matchCase,
        limit: FIND_LIMIT,
      }));
    if (!result.ok) {
      setPhase({ kind: 'refused', message: refusalMessage(result) });
      return null;
    }
    const sheetOrder = ranges.map((range) => range.sheetId);
    const matches = [...result.matches]
      .sort((a, b) => compareCells(sheetOrder, a.cell, b.cell));
    cache.current = { key, matches, sheetOrder, truncated: result.truncated };
    return cache.current;
  };

  const selectMatch = async (found: FindCache, index: number) => {
    cursor.current = found.matches[index].cell;
    setMatch({ index, total: found.matches.length, truncated: found.truncated });
    await Promise.resolve(api.reveal?.(found.matches[index].cell));
  };

  const stepMatch = async (direction: 1 | -1) => {
    const query = inputRef.current?.value ?? '';
    if (!target || query.length === 0 || busy || readOnly) return;
    setPhase({ kind: 'busy' });
    try {
      const found = await findMatches(query);
      if (!found) return;
      if (found.matches.length === 0) {
        setMatch({ index: -1, total: 0, truncated: false });
        setPhase({ kind: 'idle' });
        return;
      }
      setPhase({ kind: 'idle' });
      await selectMatch(found, stepFrom(found, cursor.current ?? activeCellOf(target), direction));
    } catch (error) {
      setPhase({ kind: 'refused', message: error instanceof Error ? error.message : String(error) });
    }
  };

  const replaceCurrent = async () => {
    const query = inputRef.current?.value ?? '';
    const replacement = replaceRef.current?.value ?? '';
    if (!target || query.length === 0 || busy || readOnly) return;
    setPhase({ kind: 'busy' });
    try {
      const found = await findMatches(query, REPLACE_LOOK_IN);
      if (!found) return;
      if (found.matches.length === 0) {
        setMatch({ index: -1, total: 0, truncated: false });
        setPhase({ kind: 'idle' });
        return;
      }
      const from = cursor.current ?? activeCellOf(target);
      const index = found.matches.findIndex((entry) => compareCells(found.sheetOrder, entry.cell, from) === 0);
      if (index < 0) {
        setPhase({ kind: 'idle' });
        await selectMatch(found, stepFrom(found, from, 1));
        return;
      }
      const cell = found.matches[index].cell;
      const at = { row: cell.row, col: cell.col };
      const version = await Promise.resolve(api.version());
      const request: XlsxEditRequest = {
        expectVersion: version,
        steps: [{
          op: 'replaceText',
          target: { sheetId: cell.sheetId, range: { kind: 'rowCol', start: at, end: at } },
          query,
          replacement,
          entireCell,
          lookIn: REPLACE_LOOK_IN,
          matchCase,
        }],
      };
      const validation = await Promise.resolve(api.validate(request));
      if (!validation.ok) {
        setPhase({ kind: 'refused', message: refusalMessage(validation) });
        return;
      }
      const preview = validation.previews[0];
      if (!validation.wouldApply || !preview?.wouldChange) {
        setPhase({ kind: 'preview', preview: preview ?? {
          stepIndex: 0, target: request.steps[0].target, wouldChange: false,
          changedCellCount: 0, replace: undefined,
        }, stats: preview?.replace ?? null });
        await selectMatch(found, index);
        return;
      }
      const result = await Promise.resolve(api.apply(request));
      if (!result.ok) {
        setPhase({ kind: 'refused', message: refusalMessage(result) });
        return;
      }
      setPhase({ kind: 'applied', preview, stats: preview.replace ?? null });
      onApplied?.(result);
      cache.current = null;
      cursor.current = cell;
      const fresh = await findMatches(query, REPLACE_LOOK_IN);
      if (!fresh) return;
      if (fresh.matches.length === 0) {
        setMatch({ index: -1, total: 0, truncated: false });
        return;
      }
      await selectMatch(fresh, stepFrom(fresh, cell, 1));
    } catch (error) {
      setPhase({ kind: 'refused', message: error instanceof Error ? error.message : String(error) });
    }
  };

  const runAll = async (apply: boolean) => {
    const query = inputRef.current?.value ?? '';
    const replacement = replaceRef.current?.value ?? '';
    if (!target || query.length === 0 || busy || readOnly) return;
    setPhase({ kind: 'busy' });
    try {
      const version = await Promise.resolve(api.version());
      const steps: XlsxEditStep[] = scopeTargets(target, scope).map((range) => ({
        op: 'replaceText',
        target: range,
        query,
        replacement,
        entireCell,
        lookIn: REPLACE_LOOK_IN,
        matchCase,
        ...(scope === 'workbook' ? { within: 'workbook' as const } : {}),
      }));
      if (steps.length === 0) {
        setMatch({ index: -1, total: 0, truncated: false });
        setPhase({ kind: 'idle' });
        return;
      }
      const request: XlsxEditRequest = { expectVersion: version, steps };
      const validation = await Promise.resolve(api.validate(request));
      if (!validation.ok) {
        setPhase({ kind: 'refused', message: refusalMessage(validation) });
        return;
      }
      const previews = validation.previews;
      const preview: XlsxEditPreview = previews[0] ?? {
        stepIndex: 0, target: steps[0].target, wouldChange: false,
        changedCellCount: 0, replace: undefined,
      };
      const stats = sumReplaceStats(previews, target.sheets);
      const wouldChange = validation.wouldApply && previews.some((entry) => entry.wouldChange);
      if (!wouldChange) {
        setPhase({ kind: 'preview', preview, stats });
        return;
      }
      if (!apply) {
        setPhase({ kind: 'preview', preview, stats });
        return;
      }
      const result = await Promise.resolve(api.apply(request));
      if (!result.ok) {
        setPhase({ kind: 'refused', message: refusalMessage(result) });
        return;
      }
      setPhase({ kind: 'applied', preview, stats });
      setMatch(null);
      cache.current = null;
      onApplied?.(result);
    } catch (error) {
      setPhase({ kind: 'refused', message: error instanceof Error ? error.message : String(error) });
    }
  };

  const onKeyDown = (event: ReactKeyboardEvent) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onClose();
    }
  };

  const beginDrag = (event: ReactMouseEvent) => {
    if (event.button !== 0) return;
    if ((event.target as HTMLElement).closest('button, input, label')) return;
    event.preventDefault();
    dragOrigin.current = {
      x: event.clientX,
      y: event.clientY,
      baseX: position.x,
      baseY: position.y,
    };
    setDragging(true);
  };

  return (
    <div
      style={{
        position: 'fixed', inset: 0, zIndex: 10001, display: 'flex',
        alignItems: 'flex-start', justifyContent: 'flex-end', padding: 16,
        pointerEvents: 'none',
      }}
    >
      <form
        role="dialog"
        aria-modal="false"
        aria-labelledby={titleId}
        onSubmit={(event) => { event.preventDefault(); void replaceCurrent(); }}
        onKeyDown={onKeyDown}
        style={{
          width: `min(${DIALOG_WIDTH}px, calc(100vw - 32px))`,
          padding: 16, borderRadius: 8, background: '#ffffff', color: toolbarTheme.ink,
          boxShadow: '0 12px 32px rgba(0,0,0,.24)', fontSize: 13,
          display: 'flex', flexDirection: 'column', gap: 10,
          pointerEvents: 'auto',
          transform: `translate(${position.x}px, ${position.y}px)`,
        }}
      >
        <div
          onMouseDown={beginDrag}
          style={{
            display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12,
            cursor: dragging ? 'grabbing' : 'move', userSelect: 'none',
          }}
        >
          <strong id={titleId}>{t('replaceDialog.title')}</strong>
          <button type="button" onClick={onClose} aria-label={t('commands.cancel')}
            style={{ ...BUTTON, height: 26, padding: '0 8px', border: '1px solid #dadce0', background: '#fff', cursor: 'pointer' }}>
            ×
          </button>
        </div>

        <div style={{ display: 'grid', gap: 4 }}>
          <label htmlFor={findId}>{t('replaceDialog.find')}</label>
          <div style={{ display: 'flex', gap: 6 }}>
            <input id={findId} ref={inputRef}
              autoComplete="off"
              onKeyDown={(event) => {
                if (event.key !== 'Enter') return;
                event.preventDefault();
                void stepMatch(event.shiftKey ? -1 : 1);
              }}
              style={{ flex: 1, minWidth: 0, height: 30, padding: '0 8px', border: '1px solid #dadce0', borderRadius: 4 }} />
            <button type="button" disabled={!ready} onClick={() => void stepMatch(-1)}
              aria-label={t('replaceDialog.findPrevious')} title={t('replaceDialog.findPrevious')}
              style={{ ...NAV_BUTTON, border: '1px solid #dadce0', background: '#fff', cursor: 'pointer' }}>
              ↑
            </button>
            <button type="button" disabled={!ready} onClick={() => void stepMatch(1)}
              aria-label={t('replaceDialog.findNext')} title={t('replaceDialog.findNext')}
              style={{ ...NAV_BUTTON, border: '1px solid #dadce0', background: '#fff', cursor: 'pointer' }}>
              ↓
            </button>
          </div>
        </div>

        <div style={{ display: 'grid', gap: 4 }}>
          <label htmlFor={replaceId}>{t('replaceDialog.replaceWith')}</label>
          <div style={{ display: 'flex', gap: 6 }}>
            <input id={replaceId} ref={replaceRef}
              autoComplete="off"
              onKeyDown={(event) => {
                if (event.key !== 'Enter') return;
                event.preventDefault();
                void replaceCurrent();
              }}
              style={{ flex: 1, minWidth: 0, height: 30, padding: '0 8px', border: '1px solid #dadce0', borderRadius: 4 }} />
            <button type="button" disabled={!ready} onClick={() => void replaceCurrent()}
              style={{ ...BUTTON, border: '1px solid #dadce0', background: '#fff', cursor: 'pointer' }}>
              {t('replaceDialog.replace')}
            </button>
          </div>
        </div>

        <fieldset style={{ border: 'none', margin: 0, padding: 0, display: 'grid', gap: 6 }}>
          <legend id={lookInId} style={{ padding: 0, marginBottom: 4 }}>{t('replaceDialog.lookIn')}</legend>
          <div role="radiogroup" aria-labelledby={lookInId} style={{ display: 'flex', gap: 16 }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="radio" name="replace-look-in" checked={lookIn === 'values'}
                onChange={() => { cache.current = null; setMatch(null); setLookIn('values'); }} /> {t('replaceDialog.values')}
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="radio" name="replace-look-in" checked={lookIn === 'formulas'}
                onChange={() => { cache.current = null; setMatch(null); setLookIn('formulas'); }} /> {t('replaceDialog.formulas')}
            </label>
          </div>
        </fieldset>

        <label htmlFor={entireId} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <input id={entireId} type="checkbox" checked={entireCell}
            onChange={(event) => {
              cache.current = null;
              setMatch(null);
              setEntireCell(event.target.checked);
            }} />
          {t('replaceDialog.entireCell')}
        </label>

        <label htmlFor={matchCaseId} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <input id={matchCaseId} type="checkbox" checked={matchCase}
            onChange={(event) => {
              cache.current = null;
              setMatch(null);
              setMatchCase(event.target.checked);
            }} />
          {t('replaceDialog.matchCase')}
        </label>

        <label htmlFor={scopeId} style={{ display: 'grid', gap: 4 }}>
          {t('replaceDialog.scope')}
          <select
            id={scopeId}
            value={scope}
            onChange={(event) => {
              cache.current = null;
              setMatch(null);
              setPhase({ kind: 'idle' });
              setScope(event.target.value as ReplaceScope);
            }}
            style={{ height: 30, padding: '0 8px', border: '1px solid #dadce0', borderRadius: 4, background: '#fff' }}
          >
            <option value="selection">{t('replaceDialog.scopeSelection')}</option>
            <option value="sheet">{t('replaceDialog.scopeSheet')}</option>
            <option value="workbook">{t('replaceDialog.scopeWorkbook')}</option>
          </select>
        </label>

        <div style={{ fontSize: 12, color: toolbarTheme.inkMuted }}>
          {target
            ? scope === 'selection'
              ? t('replaceDialog.range', { range: target.label })
              : scope === 'sheet'
                ? t('replaceDialog.sheetScope', { sheet: activeSheetName })
                : t('replaceDialog.workbookScope', { count: workbookSheetCount })
            : t('replaceDialog.noRange')}
        </div>
        <div style={{ fontSize: 12, color: '#b06000' }}>{t('replaceDialog.formulaWarning')}</div>

        <div role="status" aria-live="polite" style={{ minHeight: 18, fontSize: 12, color: toolbarTheme.inkMuted }}>
          {phase.kind === 'refused' && <span style={{ color: '#b3261e' }}>{t('replaceDialog.refused', { message: phase.message })}</span>}
          {phase.kind === 'busy' && t('replaceDialog.working')}
          {phase.kind !== 'busy' && match?.total === 0 && t('replaceDialog.noMatches')}
          {match && match.total > 0 && t(
            match.truncated ? 'replaceDialog.matchPositionTruncated' : 'replaceDialog.matchPosition',
            { index: match.index + 1, total: new Intl.NumberFormat(resolvedLocale(lang)).format(match.total) },
          )}
          {phase.kind === 'preview' && !phase.preview.wouldChange && !match && !stats?.matched &&
            t('replaceDialog.noMatches')}
          {stats?.matched
            ? ` ${t('replaceDialog.summary', {
              matched: stats.matched, replaced: stats.replaced,
              formulas: stats.skippedFormulas, locked: stats.skippedLocked,
            })}`
            : null}
          {stats?.skippedValues
            ? ` ${t('replaceDialog.skippedValues', {
              count: stats.skippedValues, cells: skippedCellList(stats),
            })}`
            : null}
          {stats && refusedTotal(stats) > 0
            ? <span style={{ color: '#b3261e' }}>
              {` ${t('replaceDialog.refused', { message: refusedCellList(stats) })}`}
            </span>
            : null}
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
          <button type="button" disabled={!ready} onClick={() => void runAll(false)}
            style={{ ...BUTTON, border: '1px solid #dadce0', background: '#fff', cursor: 'pointer' }}>
            {t('replaceDialog.preview')}
          </button>
          <button type="button" disabled={!ready} onClick={() => void runAll(true)}
            style={{ ...BUTTON, border: '1px solid #217346', background: '#217346', color: '#fff', cursor: 'pointer' }}>
            {t('replaceDialog.replaceAll')}
          </button>
        </div>
      </form>
    </div>
  );
}
