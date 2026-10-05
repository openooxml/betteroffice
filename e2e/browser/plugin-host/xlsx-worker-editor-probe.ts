export interface WorkerEditorProbe {
  ready: Promise<void>;
  peerEntries: { generation: number; method: string; args: unknown[] }[];
  replayEntries: { generation: number; sequence: number; method: string; args: unknown[] }[];
  generation(): number;
  releaseHydration(): void;
  flush(): Promise<void>;
  zoom(scale: number): Promise<void>;
  cellClip(): { x: number; y: number; width: number; height: number };
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
  hydrated(): boolean;
  holdPreview(): void;
  previewHeld(): boolean;
  releasePreview(text: string): void;
  paintedTexts(): string[];
  cell(): Promise<string | null>;
  saveAndReopen(): Promise<number>;
  undo(): Promise<boolean>;
}

declare global {
  interface Window { __xlsxWorkerEditor: WorkerEditorProbe }
}
