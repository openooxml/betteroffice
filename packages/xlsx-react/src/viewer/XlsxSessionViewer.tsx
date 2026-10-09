import {
  buildA11yGrid, cellAtPoint, chartRegionAtPoint, extendTo, hyperlinkAtCell,
  normalizeRange, paintDisplayList, parseHyperlinkLocation, rangeRect,
  safeExternalHyperlink, selectionAt, selectionKeyReducer, toTsv,
} from '@betteroffice/xlsx';
import type { CellAddr, CellEdit, Selection } from '@betteroffice/xlsx';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, MouseEvent } from 'react';
import type { XlsxWorkerViewerProps } from '../XlsxEditor';
import { createXlsxCommandController } from '../commands/createXlsxCommandStore';
import type { XlsxCommandBinding } from '../commands/createXlsxCommandStore';
import { commandForEvent } from '../commands/descriptors';
import { DEFAULT_FONT_FAMILIES, DEFAULT_FONT_SIZES } from '../commands/evaluate';
import type { XlsxCommandEnvironment } from '../commands/evaluate';
import type { XlsxCommandArgs, XlsxCommandResult } from '../commands/types';
import { useCommandShortcuts } from '../commands/useCommandShortcuts';
import { XlsxCommandContext } from '../commands/XlsxCommandProvider';
import { EditorChromeContext } from '../components/EditorToolbarContext';
import { FormulaBarContext } from '../components/toolbar/FormulaBar';
import type { FormulaBarBinding } from '../components/toolbar/FormulaBar';
import { useTranslation } from '../i18n';
import { deriveLimits, expandRangeToMergedCells, scaledRect } from './sessionGeometry';
import { useSessionWorkbook } from './useSessionWorkbook';
import type { ViewerSession, ViewerSurface } from './useSessionWorkbook';

const BRAND = '#217346';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const visuallyHidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, border: 0,
  overflow: 'hidden', clip: 'rect(0 0 0 0)', whiteSpace: 'nowrap',
};

function address(cell: CellAddr): string {
  let column = '';
  for (let col = cell.col + 1; col > 0; col = Math.floor((col - 1) / 26)) {
    column = String.fromCharCode(65 + (col - 1) % 26) + column;
  }
  return `${column}${cell.row + 1}`;
}

function selectedRange(selection: Selection): string {
  const range = normalizeRange(selection);
  return `${address({ row: range.top, col: range.left })}:${address({ row: range.bottom, col: range.right })}`;
}

function download(bytes: Uint8Array, name: string): void {
  const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: XLSX_MIME }));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

export function XlsxSessionViewer(props: XlsxWorkerViewerProps) {
  const { t } = useTranslation();
  const rootRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [commands] = useState(createXlsxCommandController);
  const [zoom, setZoom] = useState(1);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  const [selectedChart, setSelectedChart] = useState<string | null>(null);
  const [focusedCell, setFocusedCell] = useState<CellEdit | null>(null);
  const dragging = useRef(false);
  const clickStart = useRef<CellAddr | null>(null);
  const focus = useCallback(() => scrollRef.current?.focus({ preventScroll: true }), []);
  const { run, loading, error, reportError } = useSessionWorkbook(props, commands.store, focus);
  const latest = useRef({ run, props, zoom, selectedChart, loading, t, reportError });
  latest.current = { run, props, zoom, selectedChart, loading, t, reportError };
  const pendingSave = useRef<{ run: ViewerSession; promise: Promise<XlsxCommandResult> } | null>(null);

  const surface = useMemo<ViewerSurface>(() => ({
    capture() {
      const run = latest.current.run;
      const scroll = scrollRef.current;
      const zoom = zoomRef.current;
      if (!run?.current || !run.view || !scroll || !canvasRef.current) return null;
      const width = scroll.clientWidth;
      const height = scroll.clientHeight;
      if (width === 0 || height === 0) return null;
      return {
        sheet: run.active, navigation: run.navigation, zoom,
        dpr: window.devicePixelRatio || 1, width, height,
        viewport: {
          x: scroll.scrollLeft / zoom, y: scroll.scrollTop / zoom,
          width: width / zoom, height: height / zoom,
        },
      };
    },
    paint(frame, request) {
      const canvas = canvasRef.current!;
      const context = canvas.getContext('2d');
      if (!context) throw new Error('Canvas context is unavailable');
      canvas.width = Math.round(request.width * request.dpr);
      canvas.height = Math.round(request.height * request.dpr);
      canvas.style.width = `${request.width}px`;
      canvas.style.height = `${request.height}px`;
      paintDisplayList(context, frame.displayList, request.dpr * request.zoom);
    },
  }), []);

  useLayoutEffect(() => run?.connect(surface), [run, surface]);
  useLayoutEffect(() => {
    if (!run?.current || !run.scrollTarget || !scrollRef.current) return;
    scrollRef.current.scrollLeft = run.scrollTarget.x * zoomRef.current;
    scrollRef.current.scrollTop = run.scrollTarget.y * zoomRef.current;
    dragging.current = false;
    clickStart.current = null;
    setSelectedChart(null);
    run.placedView();
  }, [run, run?.scrollRevision]);

  useEffect(() => {
    const scroll = scrollRef.current;
    if (!scroll || !run) return;
    const schedule = () => run.schedule();
    schedule();
    scroll.addEventListener('scroll', schedule, { passive: true });
    const observer = new ResizeObserver(schedule);
    observer.observe(scroll);
    return () => {
      scroll.removeEventListener('scroll', schedule);
      observer.disconnect();
    };
  }, [run, zoom]);

  useEffect(() => {
    if (typeof window.matchMedia !== 'function') return;
    let media: MediaQueryList | undefined;
    const update = () => {
      latest.current.run?.schedule();
      media?.removeEventListener('change', update);
      media = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
      media.addEventListener('change', update);
    };
    update();
    return () => media?.removeEventListener('change', update);
  }, []);

  useEffect(() => {
    const stop = () => { dragging.current = false; };
    window.addEventListener('mouseup', stop);
    return () => window.removeEventListener('mouseup', stop);
  }, []);

  const binding = useMemo<XlsxCommandBinding>(() => ({
    environment(): XlsxCommandEnvironment {
      const { run, selectedChart, loading, t } = latest.current;
      const selection = run?.selection;
      return {
        status: run?.current ? 'ready' : loading ? 'loading' : 'empty',
        readOnly: true, collaborative: false, canUndo: false, canRedo: false, pendingInput: false,
        selection: selectedChart ? 'chart' : run && selection ? {
          sheet: run.active, ...normalizeRange(selection),
          merged: run.painted?.frame.sheet === run.active ? run.painted.frame.mergedRanges ?? [] : [],
          formatting: null,
        } : null,
        zoom: zoomRef.current, paintFormat: false, borderStyle: null, borderColor: null,
        pngExport: false, proposals: null, proposalsPanelOpen: false,
        fontFamilies: DEFAULT_FONT_FAMILIES, fontSizes: DEFAULT_FONT_SIZES, translate: t,
      };
    },
    ordered: () => false,
    admit: async (operation) => operation(),
    perform(id, args) {
      const { run, props } = latest.current;
      if (!run?.current) return { ok: false, failure: { code: 'document-replaced', message: 'Workbook is no longer open' } };
      if (id === 'zoom') {
        const next = (args as XlsxCommandArgs['zoom']).scale;
        const previous = zoomRef.current;
        zoomRef.current = next;
        setZoom(next);
        run.schedule();
        return { ok: true, status: next === previous ? 'noop' : 'executed' };
      }
      if (id === 'print') return run.afterPaint().then((painted): XlsxCommandResult => {
        if (!run.current) return { ok: false, failure: { code: 'document-replaced', message: 'Workbook is no longer open' } };
        if (!painted) return { ok: false, failure: { code: 'render-failed', message: 'Workbook was not painted' } };
        window.print();
        return { ok: true, status: 'executed' };
      });
      if (id !== 'save') return { ok: false, failure: { code: 'read-only', message: 'The editor is read-only' } };
      if (pendingSave.current?.run === run) return pendingSave.current.promise;
      const promise = run.session.save().then((bytes): XlsxCommandResult => {
        if (!run.current) return { ok: false, failure: { code: 'document-replaced', message: 'Workbook is no longer open' } };
        if (props.onSave) props.onSave(bytes);
        else download(bytes, props.fileName ?? 'workbook.xlsx');
        return { ok: true, status: 'executed' };
      }).catch((error): XlsxCommandResult => {
        if (run.current) latest.current.reportError(error);
        return { ok: false, failure: {
          code: run.current ? 'command-failed' : 'document-replaced', message: String(error),
        } };
      }).finally(() => { if (pendingSave.current?.promise === promise) pendingSave.current = null; });
      pendingSave.current = { run, promise };
      return promise;
    },
    capture: () => latest.current.run ? {
      generation: latest.current.run.generation,
      target: JSON.stringify([latest.current.run.active, latest.current.run.selection]),
    } : null,
    resume: (origin) => !latest.current.run?.current || origin.generation !== latest.current.run.generation
      ? 'document-replaced' : origin.target !== JSON.stringify([latest.current.run.active, latest.current.run.selection])
        ? 'target-changed' : null,
    chrome: () => ({ i18n: latest.current.props.i18n }), focusEditor: focus,
  }), [focus]);
  useLayoutEffect(() => {
    commands.attach(binding);
    return () => commands.detach(binding);
  }, [commands, binding]);
  useLayoutEffect(() => commands.refresh());
  useCommandShortcuts(commands, rootRef);

  const painted = run?.painted;
  const frame = painted?.frame.displayList;
  const paintedZoom = painted?.request.zoom ?? zoom;
  const active = run?.active ?? 0;
  const selection = run?.selection ?? null;
  const visibleSelection = painted?.frame.sheet === active ? selection : null;
  const sheets = run?.session.state.sheets ?? [];
  const toolbarContent = props.showToolbar === false ? null : props.toolbar ?? null;

  useEffect(() => {
    if (selectedChart && frame && !frame.charts?.some((chart) => chart.id === selectedChart)) {
      setSelectedChart(null);
    }
  }, [frame, selectedChart]);

  useEffect(() => {
    setFocusedCell(null);
    if (!run?.current || !selection || !toolbarContent) return;
    let cancelled = false;
    const navigation = run.navigation;
    const a1 = address(selection.focus);
    void run.session.call.cellInputs(active, a1).then((result) => {
      if (!cancelled && run.current && run.navigation === navigation) setFocusedCell(result.cells[0]?.[0] ?? null);
    }).catch((error) => {
      if (!cancelled && run.current && run.navigation === navigation) reportError(error);
    });
    return () => { cancelled = true; };
  }, [run, selection, active, toolbarContent, reportError]);

  const formulaBar: FormulaBarBinding = {
    a1: focusedCell?.a1 ?? '', value: focusedCell?.input ?? '',
    disabled: !run?.view || !selection, readOnly: true, inputRef: null,
    onChange: () => {}, onCommit: () => {}, onCancel: () => {}, onBlur: () => {},
    onCompositionStart: () => {}, onCompositionEnd: () => {},
  };

  const limits = () => deriveLimits(frame ?? null, run!.view!,
    (scrollRef.current?.clientHeight ?? 0) / paintedZoom);
  const point = (event: MouseEvent): { cell: CellAddr | null; chart: string | null } => {
    if (!frame || !canvasRef.current) return { cell: null, chart: null };
    const rect = canvasRef.current.getBoundingClientRect();
    const x = (event.clientX - rect.left) / paintedZoom;
    const y = (event.clientY - rect.top) / paintedZoom;
    return {
      cell: frame.grid ? cellAtPoint(frame.grid, x, y) : null,
      chart: chartRegionAtPoint(frame.charts, x, y)?.id ?? null,
    };
  };
  const onMouseDown = (event: MouseEvent) => {
    if (!run?.view || painted?.frame.sheet !== active) return;
    const { cell, chart } = point(event);
    clickStart.current = null;
    setSelectedChart(chart);
    if (chart) { focus(); event.preventDefault(); return; }
    if (!cell || !selection) return;
    run.setSelection(event.shiftKey ? extendTo(selection, cell, limits()) : selectionAt(cell));
    clickStart.current = cell;
    dragging.current = true;
    focus();
  };
  const onMouseMove = (event: MouseEvent<HTMLDivElement>) => {
    const { cell, chart } = point(event);
    if (!cell) {
      if (!dragging.current) { event.currentTarget.style.cursor = 'default'; event.currentTarget.title = ''; }
      return;
    }
    if (!dragging.current) {
      const hyperlink = frame ? hyperlinkAtCell(frame, cell.row, cell.col) : null;
      event.currentTarget.style.cursor = chart || hyperlink ? 'pointer' : 'default';
      event.currentTarget.title = chart ? '' : hyperlink?.tooltip ?? '';
      return;
    }
    if (run?.selection && run.view && painted?.frame.sheet === active) {
      run.setSelection(extendTo(run.selection, cell, limits()));
    }
  };
  const onClick = (event: MouseEvent) => {
    const { cell } = point(event);
    const start = clickStart.current;
    clickStart.current = null;
    if (!run || !frame || !cell || !start || cell.row !== start.row || cell.col !== start.col) return;
    const hyperlink = hyperlinkAtCell(frame, cell.row, cell.col);
    if (!hyperlink) return;
    const href = safeExternalHyperlink(hyperlink);
    if (href) { window.open(href, '_blank', 'noopener,noreferrer'); return; }
    if (!hyperlink.location) return;
    const destination = parseHyperlinkLocation(hyperlink.location, sheets[painted!.frame.sheet]?.name ?? '');
    if (!destination) return;
    const sheet = sheets.findIndex((sheet) => sheet.name.toLowerCase() === destination.sheetName.toLowerCase());
    if (sheet >= 0) void run.selectCellsAsync(sheet, selectionAt(destination));
  };
  const copy = async () => {
    if (!run?.current || !selection) return;
    try {
      const inputs = await run.session.call.cellInputs(active, selectedRange(selection));
      if (run.current) await navigator.clipboard.writeText(toTsv(inputs.cells));
    } catch {}
  };
  const onKeyDown = (event: KeyboardEvent) => {
    if (!run?.view || !selection || commandForEvent(event)) return;
    if (selectedChart) {
      if (event.key === 'Escape') setSelectedChart(null);
      if (event.key === 'Escape' || event.key.startsWith('Arrow')) event.preventDefault();
      return;
    }
    if (event.ctrlKey || event.metaKey) {
      const key = event.key.toLowerCase();
      if (key === 'c' || key === 'x') { void copy(); event.preventDefault(); return; }
      if (key === 'v') { event.preventDefault(); return; }
    }
    const action = selectionKeyReducer(selection, event, limits());
    if (action.type === 'move') run.setSelection(action.selection);
    if (action.type !== 'none') event.preventDefault();
  };

  const merged = painted?.frame.mergedRanges?.slice(0, 1024) ?? [];
  const grid = frame?.grid;
  const selected = visibleSelection ? expandRangeToMergedCells(normalizeRange(visibleSelection), merged) : null;
  const focused = visibleSelection ? expandRangeToMergedCells({
    top: visibleSelection.focus.row, left: visibleSelection.focus.col,
    bottom: visibleSelection.focus.row, right: visibleSelection.focus.col,
  }, merged) : null;
  const selectionRect = grid && selected ? rangeRect(grid, selected) : null;
  const focusRect = grid && focused ? rangeRect(grid, focused) : null;
  const chartRect = selectedChart ? frame?.charts?.find((chart) => chart.id === selectedChart)?.rect : null;
  const a11yGrid = frame && painted ? buildA11yGrid(frame, visibleSelection,
    sheets[painted.frame.sheet]?.name ?? '', {
      gridLabel: t('a11y.gridLabel'), rowHeaderLabel: t('a11y.rowHeaderLabel'),
      columnHeaderLabel: t('a11y.columnHeaderLabel'), cellLabel: t('a11y.cellLabel'),
      cellLabelSelected: t('a11y.cellLabelSelected'), emptyCellLabel: t('a11y.emptyCellLabel'),
      emptyCellLabelSelected: t('a11y.emptyCellLabelSelected'),
    }) : null;
  const outline = (rect: NonNullable<typeof selectionRect>, border: string, background?: string): CSSProperties => {
    const scaled = scaledRect(rect, paintedZoom);
    return {
      position: 'absolute', left: scaled.x, top: scaled.y, width: scaled.w, height: scaled.h,
      boxSizing: 'border-box', border, background,
    };
  };

  return (
    <XlsxCommandContext.Provider value={commands.store}>
      <EditorChromeContext.Provider value={true}>
        <FormulaBarContext.Provider value={formulaBar}>
          <div ref={rootRef} className={props.className} role="application" aria-label={t('editor.appLabel')}
            style={{ position: 'relative', display: 'flex', flexDirection: 'column', width: '100%',
              height: '100%', minWidth: 0, color: '#202124', background: '#ffffff',
              fontFamily: 'ui-sans-serif, system-ui, sans-serif' }}>
            {toolbarContent != null && <div data-testid="xlsx-toolbar" style={{ flex: '0 0 auto' }}>{toolbarContent}</div>}
            <div data-testid="xlsx-workspace" style={{ position: 'relative', display: 'flex', flex: 1, minWidth: 0, minHeight: 0 }}>
              <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, minHeight: 0 }}>
                <div ref={scrollRef} data-testid="xlsx-scroll" tabIndex={0} onKeyDown={onKeyDown}
                  onMouseDown={onMouseDown} onMouseMove={onMouseMove} onClick={onClick}
                  onMouseLeave={(event) => { event.currentTarget.style.cursor = 'default'; event.currentTarget.title = ''; }}
                  style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none' }}>
                  <div style={{ position: 'absolute', top: 0, left: 0,
                    width: run?.view ? run.view.contentWidth * zoom : '100%',
                    height: run?.view ? run.view.contentHeight * zoom : '100%' }} />
                  <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
                    <canvas ref={canvasRef} style={{ display: 'block', position: 'absolute', top: 0, left: 0 }} />
                    <div data-testid="xlsx-overlay-host" style={{ position: 'absolute', top: 0, left: 0,
                      width: 0, height: 0, pointerEvents: 'none' }}>
                      {selectionRect && !selectedChart && <div data-testid="xlsx-selection"
                        style={outline(selectionRect, `1px solid ${BRAND}`, 'rgba(33, 115, 70, 0.12)')} />}
                      {focusRect && !selectedChart && <div style={outline(focusRect, `2px solid ${BRAND}`)} />}
                      {chartRect && <div data-testid="xlsx-chart-selection" data-chart-id={selectedChart} aria-hidden
                        style={{ ...outline(chartRect, `2px solid ${BRAND}`, 'transparent'), boxShadow: '0 1px 6px rgba(0, 0, 0, 0.25)' }} />}
                    </div>
                  </div>
                </div>
              </div>
            </div>
            {a11yGrid && <>
              <div style={visuallyHidden} role="grid" aria-label={a11yGrid.label}>
                <div role="row"><span role="columnheader" />{a11yGrid.columnHeaders.map((header) =>
                  <span key={header.col} role="columnheader">{header.label}</span>)}</div>
                {a11yGrid.rows.map((row) => <div key={row.row} role="row">
                  <span role="rowheader">{row.header}</span>{row.cells.map((cell) =>
                    <span key={cell.col} role="gridcell" aria-selected={cell.selected}>{cell.label}</span>)}
                </div>)}
              </div>
              {a11yGrid.charts.map((chart, index) =>
                <div key={`${index}:${chart.label}`} style={visuallyHidden} role="img" aria-label={chart.label} />)}
            </>}
            {error && <div data-testid="xlsx-error" role="alert" style={{ position: 'absolute', inset: 0,
              display: 'grid', placeItems: 'center', padding: 16, textAlign: 'center', color: '#b00020' }}>
              {t('editor.openError')}: {error}
            </div>}
            {sheets.length > 0 && <div data-testid="xlsx-sheet-tabs" role="tablist" aria-label={t('editor.sheetTabsLabel')}
              style={{ display: 'flex', gap: 2, padding: '4px 6px', borderTop: '1px solid #e0e0e0',
                background: '#fafafa', overflowX: 'auto' }}>
              {sheets.map((sheet, index) => <button key={sheet.id} role="tab" aria-selected={index === active}
                onClick={() => { setSelectedChart(null); void run?.show(index); }}
                style={{ border: 'none', padding: '4px 12px', cursor: 'pointer',
                  borderBottom: index === active ? `2px solid ${BRAND}` : '2px solid transparent',
                  fontWeight: index === active ? 600 : 400, background: index === active ? '#ffffff' : 'transparent' }}>
                {sheet.name}
              </button>)}
            </div>}
          </div>
        </FormulaBarContext.Provider>
      </EditorChromeContext.Provider>
    </XlsxCommandContext.Provider>
  );
}
