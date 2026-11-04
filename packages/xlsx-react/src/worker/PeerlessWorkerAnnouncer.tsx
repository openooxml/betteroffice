import { createContext, memo, useContext, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, RefObject } from 'react';
import type { WorkbookSheetSummary } from '@betteroffice/xlsx';
import { useTranslation } from '../i18n';
import { WorkerPresentationStore, useWorkerPresentation } from './WorkerPresentationStore';
import {
  cellAddress, isSingleCell, paintedCellText, selectionAnnouncement, selectionKey, type AnnouncementStrings,
} from './selectionAnnouncement';

/** The loaded input of the active cell, keyed `${sheet}:${A1}`. */
export const ActiveCellInputContext = createContext<{ at: string; value: string } | null>(null);

const SETTLE_MS = 150;
const WAIT_MS = 600;

const selectSelection = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.selection;
const selectPainted = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.painted;

const visuallyHidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, overflow: 'hidden',
  clipPath: 'inset(50%)', whiteSpace: 'nowrap', border: 0,
};

/**
 * Speaks the settled selection while focus is in the grid: one polite live
 * region, so rapid arrow keys collapse into the cell the user stops on.
 */
export const PeerlessWorkerAnnouncer = memo(function PeerlessWorkerAnnouncer({ store, active, sheets, focusTarget, announceRef }: {
  store: WorkerPresentationStore;
  active: number;
  sheets: WorkbookSheetSummary[];
  /** Announcements only run while focus is inside this element. */
  focusTarget: RefObject<HTMLElement | null>;
  /** Receives a callback that repeats the current announcement, for when focus enters the grid. */
  announceRef: { current: (() => void) | null };
}) {
  const { t } = useTranslation();
  const painted = useWorkerPresentation(store, selectPainted);
  const selection = useWorkerPresentation(store, selectSelection);
  const input = useContext(ActiveCellInputContext);
  const [announcement, setAnnouncement] = useState('');
  const [focusRevision, setFocusRevision] = useState(0);
  const tick = useRef(false);
  /** The key last spoken, suffixed `?` when it was spoken before its cell data arrived. */
  const spoken = useRef<string | null>(null);
  const lastSheet = useRef<number | null>(null);
  useLayoutEffect(() => {
    announceRef.current = () => { spoken.current = null; lastSheet.current = null; setFocusRevision((value) => value + 1); };
    return () => { announceRef.current = null; };
  }, [announceRef]);
  const strings = useMemo<AnnouncementStrings>(() => ({
    cellLabel: t('a11y.cellLabel'),
    emptyCellLabel: t('a11y.emptyCellLabel'),
    hasFormula: (label) => t('a11y.hasFormula', { label }),
    rangeSelected: (vars) => t('a11y.rangeSelected', vars),
    sheetActivated: (vars) => t('a11y.sheetActivated', vars),
  }), [t]);
  const frame = painted?.request.sheet === active ? painted : null;
  const key = selection ? selectionKey(active, selection) : null;
  const merged = frame?.mergedRanges;
  const single = selection ? isSingleCell(selection, merged) : false;
  const text = selection && frame ? paintedCellText(frame.displayList, selection.anchor.row, selection.anchor.col, merged) : null;
  const at = selection ? `${active}:${cellAddress(selection.anchor.row, selection.anchor.col)}` : '';
  const formula = input && input.at === at ? input.value.startsWith('=') : null;
  const ready = !single || (text !== null && formula !== null);
  const position = sheets.findIndex((entry) => entry.index === active);
  const sheet = position >= 0 ? sheets[position] : undefined;
  useEffect(() => {
    if (!selection || key === null || spoken.current === key || (spoken.current === `${key}?` && !ready)) return;
    const timer = setTimeout(() => {
      const target = focusTarget.current;
      const focused = Boolean(target && target.contains(target.ownerDocument.activeElement));
      const switched = lastSheet.current !== null && lastSheet.current !== active;
      spoken.current = ready ? key : `${key}?`;
      lastSheet.current = active;
      if (!focused) return;
      const message = selectionAnnouncement(selection, { text, formula }, strings,
        switched && sheet ? { name: sheet.name, index: position, count: sheets.length } : undefined, merged);
      tick.current = !tick.current;
      setAnnouncement(tick.current ? `${message} ` : message);
    }, ready ? SETTLE_MS : WAIT_MS);
    return () => clearTimeout(timer);
  }, [selection, key, ready, text, formula, strings, active, position, sheet, sheets.length, merged, focusTarget, focusRevision]);
  return <div data-testid="xlsx-live-region" role="status" aria-live="polite" aria-atomic="true" style={visuallyHidden}>
    {announcement}
  </div>;
});
