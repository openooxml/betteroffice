/**
 * Wasm loader for the xlsx-wasm core.
 *
 * Call {@link initWasm} once before opening workbooks. The default browser path
 * streams the packaged wasm asset; non-fetch environments can pass bytes or a
 * precompiled module. Callers never see the JSON-string boundary.
 */

import initWasmModule, { XlsxDocument } from './generated/xlsx_wasm.js';
import type { InitInput } from './generated/xlsx_wasm.js';
import type { DisplayList } from '../display-list/types';

/**
 * A scrolled window into a sheet, in content pixels from the sheet origin — the
 * argument the renderer needs to build a frame. Mirrors the Rust `Viewport`.
 */
export interface Viewport {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Chrome-facing sheet metadata: tab names, active index, and the scrollable
 * content extent of the active sheet. Mirrors the Rust `SheetInfo`.
 */
export interface SheetInfo {
  sheetNames: string[];
  activeSheet: number;
  contentWidth: number;
  contentHeight: number;
}

/**
 * Result of any mutating call: whether it changed the workbook and the
 * (possibly grown) sheet metadata. Mirrors the Rust `EditResult`.
 *
 * `changed` lists the a1 addresses (on the active sheet) of *other* cells whose
 * displayed value moved as a fallout of the edit — the dependents the recalc
 * pass recomputed. It excludes the directly edited cell(s) and is omitted by
 * cores built before the recalc engine, so treat it as additive.
 */
export interface EditResult {
  applied: boolean;
  sheetInfo: SheetInfo;
  changed?: string[];
}

/**
 * The editable view of one cell: A1 address, the exact string the user would
 * edit (formulas as `=...`, guarded literals with a leading `'`), and whether
 * that string is a formula. Mirrors the Rust `CellEdit`.
 */
export interface CellEdit {
  a1: string;
  input: string;
  isFormula: boolean;
}

/** One cell of a batch edit: target coordinates plus the raw user input. */
export interface CellInputEdit {
  row: number;
  col: number;
  input: string;
}

/**
 * One edit inside a proposal. Same shape as {@link CellInputEdit} plus the
 * sheet index — a proposal's cells carry their own sheet on the wire, so it
 * can span sheets rather than being pinned to the active one.
 */
export interface ProposalEdit extends CellInputEdit {
  sheet: number;
}

/**
 * One cell of a pending proposal: where it lands and the *formatted display
 * texts* before/after applying it. `oldText`/`newText` are what the grid shows,
 * not the raw input — the ghost decoration paints `newText` directly. Mirrors
 * the Rust proposal cell.
 */
export interface ProposalCell {
  sheet: number;
  row: number;
  col: number;
  a1: string;
  oldText: string;
  newText: string;
}

/**
 * A pending, un-applied set of cell changes an agent has staged for human
 * review. `note` is the agent's rationale (may be null); `cells` are the
 * per-cell before/after previews. Applying it is one undo step.
 */
export interface Proposal {
  id: string;
  agentId: string;
  note: string | null;
  cells: ProposalCell[];
}

/**
 * Thrown by {@link WorkbookHandle.acceptProposal} when the workbook changed
 * under a proposal since it was staged (an edit touched one of its base cells)
 * and `force` was not set. `cells` are the a1 addresses that moved, so the UI
 * can name them and offer a force-apply.
 */
export class StaleProposalError extends Error {
  readonly cells: string[];
  constructor(cells: string[]) {
    super(`stale: ${cells.join(', ')}`);
    this.name = 'StaleProposalError';
    this.cells = cells;
  }
}

// the wasm signals a stale accept with a string starting `"stale: "` followed
// by a comma-separated a1 list; parse it back into the typed error.
const STALE_PREFIX = 'stale: ';

function staleErrorFrom(message: string): StaleProposalError {
  const cells = message
    .slice(STALE_PREFIX.length)
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  return new StaleProposalError(cells);
}

/**
 * A typed handle over an open workbook, hiding the wasm object and its JSON
 * boundary. Call {@link WorkbookHandle.dispose} to free the wasm memory.
 */
export interface WorkbookHandle {
  sheetInfo(): SheetInfo;
  displayList(viewport: Viewport): DisplayList;
  setActiveSheet(index: number): void;
  /**
   * apply one user input to a cell (parses number/bool/formula/text) and
   * recalc its transitive dependents. one undo step; moved dependents come back
   * in `EditResult.changed`.
   */
  editCell(sheet: number, row: number, col: number, input: string): EditResult;
  /** apply a batch of inputs (paste path) as one undo step; dependents recalc. */
  editCells(sheet: number, edits: CellInputEdit[]): EditResult;
  /** raw op-list escape hatch for structural ops (insert/delete rows, merges…). */
  applyOps(ops: unknown[]): EditResult;
  undo(): EditResult;
  redo(): EditResult;
  /** the editable view of one cell (formula bar / in-cell editor prefill). */
  cell(sheet: number, row: number, col: number): CellEdit;
  /** row-major editable views for a range, e.g. "A1:C3" (clipboard copy). */
  rangeCells(sheet: number, range: string): CellEdit[][];
  /**
   * render the current sheet viewport to png bytes via the native raster
   * backend — the same display list the canvas paints, rasterized server-side.
   * throws if the embedded wasm was built without png export (the `raster`
   * cargo feature); guard with {@link isPngExportAvailable}.
   */
  renderPng(viewport: Viewport): Uint8Array;
  /**
   * render an a1 range (default: the used range) at an optional scale to png.
   * dimensions are capped wasm-side; same availability guard as renderPng.
   */
  renderRangePng(opts?: { range?: string; scale?: number }): Uint8Array;
  /** serialize the workbook back to .xlsx bytes. */
  save(): Uint8Array;
  /**
   * stage an agent's proposed edits for human review without applying them.
   * returns the {@link Proposal} with per-cell formatted before/after previews.
   * throws if the embedded wasm predates proposals; guard with
   * {@link WorkbookHandle.isProposalsAvailable}.
   */
  propose(agentId: string, note: string | null, edits: ProposalEdit[]): Proposal;
  /** every pending proposal, oldest first. empty when none or unsupported. */
  listProposals(): Proposal[];
  /**
   * apply a pending proposal as one undo step and recalc dependents. throws
   * {@link StaleProposalError} when the base changed since it was staged and
   * `force` is not set; pass `{ force: true }` to apply anyway.
   */
  acceptProposal(id: string, opts?: { force?: boolean }): EditResult & { proposalId: string };
  /** drop a pending proposal without applying it; false when the id was unknown. */
  rejectProposal(id: string): boolean;
  /** whether the embedded wasm core was built with the proposals api. */
  isProposalsAvailable(): boolean;
  dispose(): void;
}

let initialized = false;
let initialization: Promise<void> | undefined;

export type WasmInitInput = InitInput | Promise<InitInput>;

/** Initialize the workbook engine. Concurrent calls share the same attempt. */
export function initWasm(
  input: WasmInitInput = new URL('./generated/xlsx_wasm_bg.wasm', import.meta.url)
): Promise<void> {
  if (initialized) return Promise.resolve();
  if (initialization) return initialization;
  initialization = initWasmModule({ module_or_path: input }).then(
    () => {
      initialized = true;
    },
    (error: unknown) => {
      initialization = undefined;
      throw toError(error);
    }
  );
  return initialization;
}

function requireInitialized(): void {
  if (!initialized) throw new Error('xlsx wasm is not initialized; call initWasm() first');
}

// wasm rejects throw strings; normalize them (and anything else) to Error.
function toError(e: unknown): Error {
  if (e instanceof Error) return e;
  return new Error(typeof e === 'string' ? e : String(e));
}

/**
 * Whether this environment exposes the WebAssembly runtime required by the core.
 */
export function isWasmAvailable(): boolean {
  return typeof WebAssembly === 'object';
}

/**
 * Whether the embedded wasm core was built with png export (the `raster` cargo
 * feature). The wasm-bindgen method only exists on the class when compiled in,
 * so chrome can disable an export control instead of calling and catching.
 */
export function isPngExportAvailable(): boolean {
  return typeof (XlsxDocument.prototype as { renderPng?: unknown }).renderPng === 'function';
}

/**
 * Whether the embedded wasm core exposes the proposals api (propose / accept /
 * reject / list). The methods only exist on the class when compiled in, so the
 * UI can hide proposal chrome and degrade gracefully against an older module.
 */
export function isProposalsAvailable(): boolean {
  return typeof (XlsxDocument.prototype as { proposeJson?: unknown }).proposeJson === 'function';
}

/**
 * Open a workbook from raw `.xlsx` bytes, initializing the core if needed.
 * Throws an `Error` if the bytes are not a readable workbook.
 */
export function openWorkbook(bytes: Uint8Array): WorkbookHandle {
  requireInitialized();
  let doc: XlsxDocument;
  try {
    doc = XlsxDocument.open(bytes);
  } catch (e) {
    throw toError(e);
  }
  return {
    sheetInfo(): SheetInfo {
      try {
        return JSON.parse(doc.sheetInfoJson()) as SheetInfo;
      } catch (e) {
        throw toError(e);
      }
    },
    displayList(viewport: Viewport): DisplayList {
      try {
        return JSON.parse(doc.displayListJson(JSON.stringify(viewport))) as DisplayList;
      } catch (e) {
        throw toError(e);
      }
    },
    setActiveSheet(index: number): void {
      try {
        doc.setActiveSheet(index);
      } catch (e) {
        throw toError(e);
      }
    },
    editCell(sheet: number, row: number, col: number, input: string): EditResult {
      return editResult(() => doc.editCellJson(JSON.stringify({ sheet, row, col, input })));
    },
    editCells(sheet: number, edits: CellInputEdit[]): EditResult {
      return editResult(() => doc.editCellsJson(JSON.stringify({ sheet, edits })));
    },
    applyOps(ops: unknown[]): EditResult {
      return editResult(() => doc.applyOpsJson(JSON.stringify({ ops })));
    },
    undo(): EditResult {
      return editResult(() => doc.undoJson());
    },
    redo(): EditResult {
      return editResult(() => doc.redoJson());
    },
    cell(sheet: number, row: number, col: number): CellEdit {
      return call(() => doc.cellJson(JSON.stringify({ sheet, row, col })));
    },
    rangeCells(sheet: number, range: string): CellEdit[][] {
      const parsed = call<{ cells: CellEdit[][] }>(() =>
        doc.rangeCellsJson(JSON.stringify({ sheet, range }))
      );
      return parsed.cells;
    },
    renderPng(viewport: Viewport): Uint8Array {
      const fn = (doc as { renderPng?: (v: string) => Uint8Array }).renderPng;
      if (typeof fn !== 'function') throw new Error('png export not in this build');
      try {
        return fn.call(doc, JSON.stringify(viewport));
      } catch (e) {
        throw toError(e);
      }
    },
    renderRangePng(opts?: { range?: string; scale?: number }): Uint8Array {
      const fn = (doc as { renderRangePng?: (a: string) => Uint8Array }).renderRangePng;
      if (typeof fn !== 'function') throw new Error('png export not in this build');
      try {
        return fn.call(doc, JSON.stringify(opts ?? {}));
      } catch (e) {
        throw toError(e);
      }
    },
    save(): Uint8Array {
      try {
        return doc.saveBytes();
      } catch (e) {
        throw toError(e);
      }
    },
    propose(agentId: string, note: string | null, edits: ProposalEdit[]): Proposal {
      const fn = (doc as { proposeJson?: (a: string) => string }).proposeJson;
      if (typeof fn !== 'function') throw new Error('proposals not in this build');
      return call<Proposal>(() => fn.call(doc, JSON.stringify({ agentId, note, edits })));
    },
    listProposals(): Proposal[] {
      const fn = (doc as { listProposalsJson?: () => string }).listProposalsJson;
      if (typeof fn !== 'function') return [];
      return call<{ proposals: Proposal[] }>(() => fn.call(doc)).proposals;
    },
    acceptProposal(id: string, opts?: { force?: boolean }): EditResult & { proposalId: string } {
      const fn = (doc as { acceptProposalJson?: (a: string) => string }).acceptProposalJson;
      if (typeof fn !== 'function') throw new Error('proposals not in this build');
      try {
        return JSON.parse(
          fn.call(doc, JSON.stringify({ id, force: opts?.force ?? false }))
        ) as EditResult & {
          proposalId: string;
        };
      } catch (e) {
        const message = e instanceof Error ? e.message : typeof e === 'string' ? e : String(e);
        if (message.startsWith(STALE_PREFIX)) throw staleErrorFrom(message);
        throw toError(e);
      }
    },
    rejectProposal(id: string): boolean {
      const fn = (doc as { rejectProposalJson?: (a: string) => string }).rejectProposalJson;
      if (typeof fn !== 'function') return false;
      return call<{ removed: boolean }>(() => fn.call(doc, JSON.stringify({ id }))).removed;
    },
    isProposalsAvailable(): boolean {
      return typeof (doc as { proposeJson?: unknown }).proposeJson === 'function';
    },
    dispose(): void {
      doc.free();
    },
  };
}

// run a wasm call returning a json string and parse it, normalizing throws.
function call<T>(fn: () => string): T {
  try {
    return JSON.parse(fn()) as T;
  } catch (e) {
    throw toError(e);
  }
}

// every mutation returns the same `{applied, sheetInfo}` envelope.
function editResult(fn: () => string): EditResult {
  return call<EditResult>(fn);
}

/**
 * The crate version string, for asserting wasm/js parity.
 */
export function wasmVersion(): string {
  requireInitialized();
  return XlsxDocument.version();
}
