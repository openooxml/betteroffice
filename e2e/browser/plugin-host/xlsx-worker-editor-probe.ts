import type { Proposal, ProposalEdit, WorkbookHandle } from '../../../packages/xlsx/src/wasm/loader';

type CommandResult = { ok: true } | { ok: false; failure: { code: string; message: string } };

/** The part of XlsxAsyncEditorApi the specs drive, typed without the built xlsx-react package. */
export interface WorkerEditorApiProbe {
  version(): Promise<string>;
  readCells(request: Parameters<WorkbookHandle['readCells']>[0]): Promise<ReturnType<WorkbookHandle['readCells']>>;
  findText(request: Parameters<WorkbookHandle['findText']>[0]): Promise<ReturnType<WorkbookHandle['findText']>>;
  validateEdits(request: Parameters<WorkbookHandle['validateEdits']>[0]): Promise<ReturnType<WorkbookHandle['validateEdits']>>;
  applyEdits(request: Parameters<WorkbookHandle['applyEdits']>[0]): Promise<ReturnType<WorkbookHandle['applyEdits']>>;
  save(): Promise<Uint8Array>;
  readonly commands: {
    getState(id: 'proposalsPanel'): { value: number };
    execute(id: 'proposalsPanel', args: { open: boolean } | null): Promise<CommandResult>;
    execute(id: 'proposalAccept' | 'proposalReject', args: { proposalId: string }): Promise<CommandResult>;
  };
  proposeChanges?(request: { agentId: string; note?: string | null; edits: ProposalEdit[] }):
    Promise<{ ok: true; proposal: Proposal } | Extract<CommandResult, { ok: false }>>;
  exportStructured(...args: Parameters<WorkbookHandle['exportStructured']>): Promise<ReturnType<WorkbookHandle['exportStructured']>>;
}

export interface WorkerEditorProbe {
  ready: Promise<void>;
  mainThreadEngines(): number;
  peerless: {
    readonly editor: WorkerEditorApiProbe;
    editorAt(index?: number): WorkerEditorApiProbe;
    cells(sheet: number, range: string, index?: number): Promise<string[][]>;
    selection(index?: number): Promise<{
      sheetIndex: number;
      cells: { anchor: { row: number; col: number }; focus: { row: number; col: number } } | null;
    } | null>;
    failure(index?: number): string | null;
    /** Cell rects of the painted frame in canvas CSS pixels, read through plugin geometry. */
    cellRects(cells: [number, number][], index?: number): Promise<{
      zoom: number;
      rects: { x: number; y: number; width: number; height: number }[];
    } | null>;
    formulaBar(index?: number): string;
    workerCount(): number;
    saveAndReopen(): Promise<number>;
    recoverAndReopen(): Promise<number>;
  } | null;
  peerEntries: { generation: number; method: string; args: unknown[] }[];
  replayEntries: { generation: number; sequence: number; method: string; args: unknown[] }[];
  generation(index?: number): number;
  readyCount(): number;
  saves: number[][];
  setReadOnly(value: boolean, toolbar?: boolean): void;
  readOnlyEditRefusal(): string;
  savedInputs(): Promise<string[]>;
  releaseHydration(): void;
  holdCellInputs(): void;
  heldCellInputs(): number;
  releaseCellInputs(): void;
  flush(): Promise<void>;
  zoom(scale: number): Promise<void>;
  cellClip(): { x: number; y: number; width: number; height: number };
  mergedGridPoints(): {
    background: { x: number; y: number };
    interiors: { x: number; y: number }[];
    edges: { x: number; y: number }[];
    neighbour: { x: number; y: number };
  };
  chartClip(): { x: number; y: number; width: number; height: number };
  queueNavigation(): void;
  adoptedSequence(): number;
  hostEdit(value: string): Promise<void>;
  queueHostEdit(value: string): void;
  queueCellHostEdit(value: string): void;
  queueStyledBatch(): void;
  queueBulkFill(): void;
  bulkFill(): Promise<void>;
  formatCells(): Promise<void>;
  pluginRoutes(): Promise<void>;
  proposals(): Promise<void>;
  history(): Promise<void>;
  fail(): void;
  recover(): Promise<number>;
  replace(): Promise<void>;
  dispose(): void;
  errors: string[];
  previews: string[];
  previewFrames: string[];
  commitOrder: { kind: 'painted-preview' | 'mutator-entry'; text: string }[];
  hydrated(index?: number): boolean;
  holdPreview(): void;
  previewHeld(): boolean;
  releasePreview(text: string): void;
  paintedTexts(index?: number): string[];
  formulaBar(index?: number): string;
  cell(index?: number): Promise<string | null>;
  saveAndReopen(): Promise<number>;
  undo(index?: number): Promise<boolean>;
  pauseDelivery(): void;
  resumeDelivery(): Promise<void>;
  pendingUpdates(sender: number): number;
  savedBytes(index?: number): number[] | null;
}

declare global {
  var __xlsxMainEngineProbe: { capture(): void } | undefined;
  interface Window { __xlsxWorkerEditor: WorkerEditorProbe }
}
