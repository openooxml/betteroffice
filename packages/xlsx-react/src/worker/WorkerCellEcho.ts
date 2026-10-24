import {
  cellRect, paintDisplayList, rangeRect, type CellInputEdit, type DisplayList, type DrawCmd, type Rect, type TextCmd,
} from '@betteroffice/xlsx';
import type { WorkbookAck } from '@betteroffice/xlsx/worker';
import type { PeerlessWorkerPaintSource as WorkerPaintSource, WorkerPaintResult } from './PeerlessWorkerPaintSource';

export interface WorkerCellEchoCell {
  sheet: number;
  row: number;
  col: number;
  text: string;
}

export interface WorkerCellEchoRange {
  sheet: number;
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface WorkerCellEchoOptions {
  source: WorkerPaintSource;
  context(): CanvasRenderingContext2D | null;
  maxPatches?: number;
  maxTextLength?: number;
}

interface CellAppearance {
  bounds: Rect;
  text: TextCmd;
}

function key(cell: WorkerCellEchoCell): string {
  return `${cell.sheet}:${cell.row}:${cell.col}`;
}

function intersection(left: Rect, right: Rect): Rect | null {
  const x = Math.max(left.x, right.x);
  const y = Math.max(left.y, right.y);
  const w = Math.min(left.x + left.w, right.x + right.w) - x;
  const h = Math.min(left.y + left.h, right.y + right.h) - y;
  return w > 0 && h > 0 ? { x, y, w, h } : null;
}

function belongs(text: TextCmd, bounds: Rect): boolean {
  return !text.chart && text.x >= bounds.x && text.x < bounds.x + bounds.w &&
    text.y >= bounds.y && text.y < bounds.y + bounds.h;
}

function subtract(bounds: Rect, covered: Rect): Rect[] {
  const overlap = intersection(bounds, covered);
  if (!overlap) return [bounds];
  const pieces = [
    { x: bounds.x, y: bounds.y, w: bounds.w, h: overlap.y - bounds.y },
    { x: bounds.x, y: overlap.y + overlap.h, w: bounds.w,
      h: bounds.y + bounds.h - overlap.y - overlap.h },
    { x: bounds.x, y: overlap.y, w: overlap.x - bounds.x, h: overlap.h },
    { x: overlap.x + overlap.w, y: overlap.y,
      w: bounds.x + bounds.w - overlap.x - overlap.w, h: overlap.h },
  ];
  return pieces.filter((rect) => rect.w > 0 && rect.h > 0);
}

export class WorkerCellEcho {
  private frame: WorkerPaintResult | null;
  private patches = new Map<number, Map<string, WorkerCellEchoCell>>();
  private appearance = new Map<string, CellAppearance>();
  private readonly maxPatches: number;
  private readonly maxTextLength: number;
  private readonly unregister: () => void;
  private disposed = false;

  constructor(private readonly options: WorkerCellEchoOptions) {
    this.maxPatches = Math.max(0, Math.floor(options.maxPatches ?? 256));
    this.maxTextLength = Math.max(0, Math.floor(options.maxTextLength ?? 4096));
    if (!Number.isSafeInteger(this.maxPatches) || !Number.isSafeInteger(this.maxTextLength)) {
      throw new RangeError('Invalid cell echo limit');
    }
    this.frame = options.source.painted;
    this.unregister = options.source.registerPaintTransaction((painted, publish) => {
      const previousFrame = this.frame;
      const previousPatches = this.patches;
      const previousAppearance = this.appearance;
      try {
        this.frame = painted;
        this.patches = new Map();
        this.appearance = new Map();
        for (const [seq, cells] of previousPatches) {
          if (painted.coversThroughSeq !== undefined && seq <= painted.coversThroughSeq) continue;
          for (const cell of cells.values()) {
            if (this.cellAppearance(cell)) this.store(seq, cell);
          }
        }
        publish();
        this.repaint(this.appearanceBounds());
      } catch (error) {
        this.frame = previousFrame;
        this.patches = previousPatches;
        this.appearance = previousAppearance;
        throw error;
      }
    });
  }

  get size(): number {
    let count = 0;
    for (const cells of this.patches.values()) count += cells.size;
    return count;
  }

  patch(seq: number, cells: Iterable<WorkerCellEchoCell>): void {
    if (this.disposed || !this.frame || this.maxPatches === 0 ||
      !Number.isSafeInteger(seq) || seq <= this.options.source.coversThroughSeq) return;
    const request = this.options.source.captureCommandFrame();
    const painted = this.frame.request;
    if (!request || request.sheet !== painted.sheet ||
      request.zoom !== painted.zoom || request.dpr !== painted.dpr ||
      request.width !== painted.width || request.height !== painted.height ||
      request.viewport.x !== painted.viewport.x || request.viewport.y !== painted.viewport.y ||
      request.viewport.width !== painted.viewport.width || request.viewport.height !== painted.viewport.height) return;
    const damaged: Rect[] = [];
    const iterator = cells[Symbol.iterator]();
    for (let inspected = 0; inspected < this.maxPatches; inspected++) {
      const next = iterator.next();
      if (next.done) break;
      const cell = next.value;
      if (!Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.col) ||
        cell.row < 0 || cell.col < 0 || cell.text.length > this.maxTextLength) continue;
      const appearance = this.cellAppearance(cell);
      if (!appearance) continue;
      damaged.push(appearance.bounds);
      this.store(seq, { ...cell });
      while (this.size > this.maxPatches) {
        const oldest = Math.min(...this.patches.keys());
        const pending = this.patches.get(oldest)!;
        const evicted = pending.keys().next().value!;
        const bounds = this.appearance.get(evicted)?.bounds;
        if (bounds) damaged.push(bounds);
        pending.delete(evicted);
        if (pending.size === 0) this.patches.delete(oldest);
      }
    }
    iterator.return?.();
    this.pruneAppearance();
    this.paint(damaged);
  }

  paste(seq: number, sheet: number, edits: readonly CellInputEdit[]): void {
    const grid = this.frame?.geometry;
    if (!grid || !this.frame || sheet !== this.frame.request.sheet || edits.length === 0) return;
    const lowerBound = (row: number, col: number) => {
      let low = 0;
      let high = edits.length;
      while (low < high) {
        const mid = Math.floor((low + high) / 2);
        const edit = edits[mid];
        if (edit.row < row || edit.row === row && edit.col < col) low = mid + 1;
        else high = mid;
      }
      return low;
    };
    const budget = this.maxPatches;
    this.patch(seq, (function* () {
      let inspected = 0;
      for (let r = 0; r < grid.rowOffsets.length - 1; r++) {
        const row = grid.rowIndices?.[r] ?? grid.startRow + r;
        const start = lowerBound(row, 0);
        const end = lowerBound(row + 1, 0);
        if (start === end) continue;
        for (let c = 0; c < grid.colOffsets.length - 1; c++) {
          const col = grid.colIndices?.[c] ?? grid.startCol + c;
          if (col < edits[start].col || col > edits[end - 1].col) continue;
          if (inspected++ >= budget) return;
          const at = lowerBound(row, col);
          const edit = edits[at];
          if (edit?.row === row && edit.col === col) yield { sheet, row, col, text: edit.input };
        }
      }
    })());
  }

  clear(seq: number, range: WorkerCellEchoRange): void {
    const grid = this.frame?.geometry;
    if (!grid || !this.frame || range.sheet !== this.frame.request.sheet) return;
    const { rowOffsets, colOffsets, rowIndices, colIndices, startRow, startCol } = grid;
    function* cells(): Iterable<WorkerCellEchoCell> {
      for (let r = 0; r < rowOffsets.length - 1; r++) {
        const row = rowIndices?.[r] ?? startRow + r;
        if (row < range.top || row > range.bottom) continue;
        for (let c = 0; c < colOffsets.length - 1; c++) {
          const col = colIndices?.[c] ?? startCol + c;
          if (col >= range.left && col <= range.right) yield { sheet: range.sheet, row, col, text: '' };
        }
      }
    }
    this.patch(seq, cells());
  }

  acknowledge(ack: Pick<WorkbookAck, 'throughSeq' | 'statuses'>): void {
    const refused = new Set<number>();
    for (const status of ack.statuses) {
      if (status.status !== 'refused' && status.status !== 'cancelled') continue;
      refused.add(status.seq);
      for (const seq of status.affectedSeqs ?? []) refused.add(seq);
    }
    this.drop(refused);
  }

  refuse(seq: number): void {
    this.drop(new Set([seq]));
  }

  dispose(): void {
    if (this.disposed) return;
    const damaged = this.appearanceBounds();
    this.patches.clear();
    this.appearance.clear();
    this.paint(damaged);
    this.disposed = true;
    this.unregister();
    this.frame = null;
  }

  private store(seq: number, cell: WorkerCellEchoCell): void {
    let pending = this.patches.get(seq);
    if (!pending) this.patches.set(seq, pending = new Map());
    pending.set(key(cell), cell);
  }

  private drop(seqs: ReadonlySet<number>): void {
    if (this.disposed) return;
    const damaged: Rect[] = [];
    for (const seq of seqs) {
      for (const cell of this.patches.get(seq)?.values() ?? []) {
        const appearance = this.appearance.get(key(cell));
        if (appearance) damaged.push(appearance.bounds);
      }
      this.patches.delete(seq);
    }
    this.pruneAppearance();
    this.paint(damaged);
  }

  private pruneAppearance(): void {
    const retained = new Set<string>();
    for (const cells of this.patches.values()) for (const id of cells.keys()) retained.add(id);
    for (const id of this.appearance.keys()) if (!retained.has(id)) this.appearance.delete(id);
  }

  private appearanceBounds(): Rect[] {
    return [...this.appearance.values()].map((cell) => cell.bounds);
  }

  private cellAppearance(cell: WorkerCellEchoCell): CellAppearance | null {
    if (!this.frame || cell.sheet !== this.frame.request.sheet) return null;
    const id = key(cell);
    const cached = this.appearance.get(id);
    if (cached) return cached;
    const rect = cellRect(this.frame.geometry, cell.row, cell.col);
    if (!rect) return null;
    const merged = this.frame.mergedRanges.find((range) => cell.row >= range.start.row &&
      cell.row <= range.end.row && cell.col >= range.start.col && cell.col <= range.end.col);
    const full = merged ? rangeRect(this.frame.geometry, { top: merged.start.row, left: merged.start.col,
      bottom: merged.end.row, right: merged.end.col }) : rect;
    if (!full) return null;
    const list = this.frame.displayList;
    const request = this.frame.request;
    const bounds = intersection(full, { x: 0, y: 0,
      w: Math.min(list.width, request.viewport.width, request.width / request.zoom),
      h: Math.min(list.height, request.viewport.height, request.height / request.zoom) });
    if (!bounds || list.charts?.some((chart) => intersection(chart.clip, bounds))) return null;
    const text = list.commands.find((command): command is TextCmd =>
      command.op === 'text' && !command.ghost && belongs(command, full));
    const appearance: CellAppearance = { bounds, text: text ? { ...text, clip: bounds } : {
      op: 'text', text: '', x: bounds.x + 4, y: bounds.y + Math.min(18, bounds.h - 2),
      fontSize: 11, color: '#000000', clip: bounds,
    } };
    this.appearance.set(id, appearance);
    return appearance;
  }

  private paint(damaged: readonly Rect[]): void {
    try { this.repaint(damaged); } catch (error) { this.options.source.fail(error); }
  }

  private repaint(damaged: readonly Rect[]): void {
    if (!this.frame || damaged.length === 0) return;
    const context = this.options.context();
    if (!context) throw new Error('Cell echo canvas context is unavailable');
    const latest = new Map<string, WorkerCellEchoCell>();
    for (const seq of [...this.patches.keys()].sort((left, right) => left - right)) {
      for (const [id, cell] of this.patches.get(seq)!) {
        latest.delete(id);
        latest.set(id, cell);
      }
    }
    const echoes: TextCmd[] = [];
    for (const [id, cell] of latest) {
      const appearance = this.appearance.get(id);
      // a number layout indexes the frame's formatted text, never the typed input
      if (appearance) echoes.push({ ...appearance.text, text: cell.text, numberFormat: undefined });
    }
    const list = this.frame.displayList;
    const viewport = { x: 0, y: 0, w: list.width, h: list.height };
    const commands: DrawCmd[] = [];
    for (const command of list.commands) {
      if (command.op !== 'text' || command.chart || echoes.length === 0) {
        commands.push(command);
        continue;
      }
      let clips = [command.clip ?? viewport];
      for (const echo of echoes) clips = clips.flatMap((clip) => subtract(clip, echo.clip!));
      for (const clip of clips) commands.push({ ...command, clip });
    }
    const displayList: DisplayList = { ...list, commands: [...commands, ...echoes] };
    const scale = this.frame.request.dpr * this.frame.request.zoom;
    context.save();
    try {
      context.setTransform(scale, 0, 0, scale, 0, 0);
      context.beginPath();
      for (const rect of damaged) context.rect(rect.x, rect.y, rect.w, rect.h);
      context.clip();
      paintDisplayList(context, displayList, scale);
    } finally { context.restore(); }
  }
}
