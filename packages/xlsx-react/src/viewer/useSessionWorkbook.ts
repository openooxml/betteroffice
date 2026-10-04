import { openWorkbookSession, rangeRect, selectionAt } from '@betteroffice/xlsx';
import type {
  Selection, Viewport, WorkbookFrame, WorkbookSession, WorkbookSheetView,
} from '@betteroffice/xlsx';
import { useCallback, useEffect, useRef, useState } from 'react';
import type { XlsxEditorProps, XlsxWorkerViewerApi, XlsxWorkerViewerProps } from '../XlsxEditor';
import type { XlsxCommandStore } from '../commands/types';
import { expandRangeToMergedCells } from './sessionGeometry';

export const workbookSessionOpener = { open: openWorkbookSession };

export interface FrameRequest {
  sheet: number;
  navigation: number;
  viewport: Viewport;
  zoom: number;
  dpr: number;
  width: number;
  height: number;
}

export interface ViewerSurface {
  capture(): FrameRequest | null;
  paint(frame: WorkbookFrame, request: FrameRequest): void;
}

function sameRequest(left: FrameRequest | null, right: FrameRequest): boolean {
  return left !== null && left.sheet === right.sheet && left.navigation === right.navigation &&
    left.zoom === right.zoom && left.dpr === right.dpr && left.width === right.width &&
    left.height === right.height && sameViewport(left.viewport, right.viewport);
}

function sameViewport(left: Viewport, right: Viewport): boolean {
  return left.x === right.x && left.y === right.y &&
    left.width === right.width && left.height === right.height;
}

function validSelection(selection: Selection): boolean {
  return [selection?.anchor, selection?.focus].every((cell) => cell &&
    Number.isInteger(cell.row) && cell.row >= 0 && cell.row <= 1_048_575 &&
    Number.isInteger(cell.col) && cell.col >= 0 && cell.col <= 16_383);
}

export class ViewerSession {
  alive = true;
  failed = false;
  active: number;
  navigation = 0;
  selection: Selection | null = selectionAt({ row: 0, col: 0 });
  view: WorkbookSheetView | null = null;
  painted: { frame: WorkbookFrame; request: FrameRequest } | null = null;
  scrollRevision = 0;
  scrollTarget: { x: number; y: number } | null = null;
  private surface: ViewerSurface | null = null;
  private raf: number | null = null;
  private revision = 0;
  private pending: { request: FrameRequest; revision: number } | null = null;
  private running = false;
  private ready = false;
  private viewRequest = 0;
  private readonly waiters = new Set<{
    request: FrameRequest | null | undefined;
    resolve: (painted: boolean) => void;
  }>();

  constructor(
    readonly session: WorkbookSession,
    private readonly changed: () => void,
    private readonly report: (error: unknown) => void,
    private readonly onReady: () => void,
    private readonly isCurrent: () => boolean = () => true,
    readonly generation = 0
  ) {
    this.active = session.state.activeSheet;
  }

  get current(): boolean { return this.alive && !this.failed && this.isCurrent(); }

  connect(surface: ViewerSurface): () => void {
    this.surface = surface;
    this.schedule();
    return () => {
      if (this.surface !== surface) return;
      this.surface = null;
      this.cancelFrame();
    };
  }

  start(): void { void this.show(this.active); }

  async show(sheet: number): Promise<boolean> {
    if (!this.current || !this.validSheet(sheet)) return false;
    this.beginNavigation();
    const viewRequest = ++this.viewRequest;
    this.active = sheet;
    this.view = null;
    this.selection = selectionAt({ row: 0, col: 0 });
    this.changed();
    try {
      const view = await this.session.call.sheetView(sheet);
      if (!this.current || viewRequest !== this.viewRequest) return false;
      this.presentView(view, { x: view.initialScrollX, y: view.initialScrollY });
      return true;
    } catch (error) {
      if (this.current && viewRequest === this.viewRequest) this.fail(error);
      return false;
    }
  }

  async selectCellsAsync(sheet: number, selection: Selection): Promise<boolean> {
    if (!this.current || !this.validSheet(sheet) || !validSelection(selection)) return false;
    const next = { anchor: { ...selection.anchor }, focus: { ...selection.focus } };
    const navigation = this.beginNavigation();
    this.viewRequest += 1;
    try {
      const [view, anchor, geometry] = await Promise.all([
        this.view?.sheet === sheet ? Promise.resolve(this.view) : this.session.call.sheetView(sheet),
        this.session.call.cellGeometry(sheet, next.anchor.row, next.anchor.col),
        this.session.call.cellGeometry(sheet, next.focus.row, next.focus.col),
      ]);
      if (!this.current || navigation !== this.navigation) return false;
      this.active = sheet;
      this.selection = next;
      const painted = new Promise<boolean>((resolve) => this.waiters.add({ request: null, resolve }));
      this.presentView({
        ...view,
        contentWidth: Math.max(view.contentWidth, anchor.rect.x + anchor.rect.w, geometry.rect.x + geometry.rect.w),
        contentHeight: Math.max(view.contentHeight, anchor.rect.y + anchor.rect.h, geometry.rect.y + geometry.rect.h),
      }, geometry.scrollPosition);
      if (!await painted || !this.current || navigation !== this.navigation) return false;
      const paintedFrame = this.painted;
      if (!paintedFrame || !sameRequest(this.surface?.capture() ?? null, paintedFrame.request)) return false;
      const { frame, request } = paintedFrame;
      const { viewport } = request;
      const focused = expandRangeToMergedCells({
        top: next.focus.row, left: next.focus.col,
        bottom: next.focus.row, right: next.focus.col,
      }, frame.mergedRanges ?? []);
      const rect = rangeRect(frame.displayList.grid, focused);
      if (!rect || rect.w <= 0 || rect.h <= 0) return false;
      const left = focused.left < view.frozenCols ? 0 : view.frozenWidth;
      const top = focused.top < view.frozenRows ? 0 : view.frozenHeight;
      const right = Math.min(viewport.width, frame.displayList.width,
        focused.right < view.frozenCols ? view.frozenWidth : Infinity);
      const bottom = Math.min(viewport.height, frame.displayList.height,
        focused.bottom < view.frozenRows ? view.frozenHeight : Infinity);
      const tolerance = 1 / request.zoom;
      return right > left && bottom > top &&
        Math.min(rect.x + rect.w, right + tolerance) > Math.max(rect.x, left - tolerance) &&
        Math.min(rect.y + rect.h, bottom + tolerance) > Math.max(rect.y, top - tolerance);
    } catch {
      if (this.current && navigation === this.navigation) {
        if (this.view) this.schedule();
        else void this.show(this.active);
      }
      return false;
    }
  }

  setSelection(selection: Selection | null): void {
    if (!this.current) return;
    this.beginNavigation();
    this.selection = selection;
    this.changed();
    this.schedule();
  }

  afterPaint(): Promise<boolean> {
    if (!this.current) return Promise.resolve(false);
    if (this.painted && sameRequest(this.surface?.capture() ?? null, this.painted.request)) {
      return Promise.resolve(true);
    }
    this.schedule();
    return new Promise((resolve) => this.waiters.add({ request: undefined, resolve }));
  }

  placedView(): void {
    this.scrollTarget = null;
    const request = this.surface?.capture() ?? null;
    for (const waiter of this.waiters) if (waiter.request === null) waiter.request = request;
    this.schedule();
  }

  schedule(): void {
    if (!this.current || !this.surface) return;
    this.revision += 1;
    if (this.raf !== null) return;
    this.raf = requestAnimationFrame(() => {
      this.raf = null;
      if (!this.current || !this.view || this.scrollTarget) return;
      const request = this.surface?.capture();
      if (!request) return;
      for (const waiter of this.waiters) if (waiter.request === null) waiter.request = request;
      this.pending = { request, revision: this.revision };
      void this.drain();
    });
  }

  fail(error: unknown): void {
    if (!this.current) return;
    this.failed = true;
    this.cancelFrame();
    this.finish(false);
    this.report(error);
    this.changed();
  }

  dispose(): void {
    if (!this.alive) return;
    this.alive = false;
    this.cancelFrame();
    this.finish(false);
    this.surface = null;
    this.painted = null;
    void this.session.dispose().catch(() => {});
  }

  private validSheet(sheet: number): boolean {
    return Number.isInteger(sheet) && sheet >= 0 && sheet < this.session.state.sheets.length;
  }

  private beginNavigation(): number {
    this.finish(false);
    this.navigation += 1;
    this.revision += 1;
    this.pending = null;
    return this.navigation;
  }

  private presentView(view: WorkbookSheetView, scroll: { x: number; y: number }): void {
    if (view.sheet !== this.active) { this.fail(new Error('Unexpected sheet view')); return; }
    this.view = view;
    this.scrollTarget = scroll;
    this.scrollRevision += 1;
    this.changed();
    this.schedule();
  }

  private cancelFrame(): void {
    if (this.raf !== null) cancelAnimationFrame(this.raf);
    this.raf = null;
    this.pending = null;
  }

  private finish(value: boolean): void {
    for (const waiter of this.waiters) waiter.resolve(value);
    this.waiters.clear();
  }

  private matches(request: FrameRequest, revision: number): boolean {
    return revision === this.revision && request.navigation === this.navigation &&
      request.sheet === this.active && this.scrollTarget === null &&
      sameRequest(this.surface?.capture() ?? null, request);
  }

  private async drain(): Promise<void> {
    if (this.running || !this.current) return;
    this.running = true;
    try {
      while (this.current && this.pending) {
        const pending = this.pending;
        this.pending = null;
        const { request, revision } = pending;
        if (!this.matches(request, revision)) {
          if (this.raf === null) this.schedule();
          continue;
        }
        try {
          const frame = await this.session.call.frame(request.viewport, { sheet: request.sheet });
          if (!this.current) break;
          if (!this.matches(request, revision)) {
            if (!this.pending && this.raf === null) this.schedule();
            continue;
          }
          if (frame.sheet !== request.sheet || !sameViewport(frame.viewport, request.viewport)) {
            throw new Error('Unexpected workbook frame');
          }
          this.surface!.paint(frame, request);
          if (!this.current) break;
          this.painted = { frame, request };
          this.changed();
          for (const waiter of this.waiters) {
            waiter.resolve(waiter.request === undefined || sameRequest(waiter.request, request));
          }
          this.waiters.clear();
          if (!this.ready) {
            this.ready = true;
            this.onReady();
          }
        } catch (error) {
          if (!this.current) break;
          if (error instanceof Error && error.name === 'SessionSuperseded') {
            if (!this.pending) this.schedule();
          } else if (this.matches(request, revision)) {
            this.fail(error);
          }
        }
      }
    } finally { this.running = false; }
  }
}

export function useSessionWorkbook(
  props: XlsxWorkerViewerProps, commands: XlsxCommandStore, focus: () => void
) {
  const latest = useRef({ props, focus });
  latest.current = { props, focus };
  const generation = useRef(0);
  const [run, setRun] = useState<ViewerSession | null>(null);
  const [, setRevision] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const collaboration = (props as XlsxWorkerViewerProps & Pick<XlsxEditorProps, 'collaboration'>).collaboration;
  const reportError = useCallback((value: unknown) => {
    const error = value instanceof Error ? value : new Error(
      value && typeof value === 'object' && 'message' in value ? String(value.message) : String(value)
    );
    setError(error.message);
    latest.current.props.onError?.(error);
  }, []);

  useEffect(() => {
    const token = ++generation.current;
    const controller = new AbortController();
    let disposed = false;
    let opened: ViewerSession | undefined;
    let offFailure = () => {};
    let cleanup: void | (() => void);
    const current = () => !disposed && token === generation.current && latest.current.props.file === props.file;
    const changed = () => { if (current()) setRevision((revision) => revision + 1); };
    setRun(null);
    setError(null);
    setLoading(Boolean(props.file));
    if (collaboration) {
      reportError(new Error('Collaboration is unavailable in the worker viewer'));
      setLoading(false);
    } else if (props.file) void (async () => {
      try {
        const session = await workbookSessionOpener.open(props.file!, { signal: controller.signal });
        if (!current()) { void session.dispose().catch(() => {}); return; }
        opened = new ViewerSession(session, changed,
          (error) => { if (current()) reportError(error); }, () => {
            if (current()) cleanup = latest.current.props.onReady?.(api);
          }, current, token);
        const viewer = opened;
        const afterSession = async <T,>(operation: (session: WorkbookSession) => Promise<T>): Promise<T | null> => {
          if (!viewer.current) return null;
          try {
            const result = await operation(session);
            return viewer.current ? result : null;
          } catch (error) {
            if (!viewer.current) return null;
            throw error;
          }
        };
        const refusal = () => afterSession(async (session) => ({
          ok: false as const, version: await session.call.version(),
          failure: { code: 'read-only' as const, message: 'The editor is read-only' },
        }));
        const api: XlsxWorkerViewerApi = {
          handle: null, commands, save: () => null, selectCells: () => false,
          refreshProposals: () => {},
          clearSelection: () => viewer.setSelection(null),
          focus: () => { if (viewer.current) latest.current.focus(); },
          selectCellsAsync: (sheet, selection) => viewer.selectCellsAsync(sheet, selection),
          saveAsync: () => afterSession((session) => session.save()),
          version: () => afterSession((session) => session.call.version()),
          readCells: (request) => afterSession((session) => session.call.readCells(request)),
          findText: (request) => afterSession((session) => session.call.findText(request)),
          validateEdits: refusal, applyEdits: refusal,
        };
        offFailure = session.onFailure((failure) => viewer.fail(failure));
        if (session.failure) viewer.fail(session.failure);
        setRun(viewer);
        setLoading(false);
        viewer.start();
      } catch (error) {
        if (!current()) return;
        opened?.fail(error);
        if (!opened) reportError(error);
        setLoading(false);
      }
    })();
    return () => {
      disposed = true;
      generation.current += 1;
      controller.abort();
      offFailure();
      try { if (typeof cleanup === 'function') cleanup(); }
      finally { opened?.dispose(); }
    };
  }, [props.file, collaboration, commands, reportError]);

  return { run: run?.current ? run : null, error, loading, reportError };
}
