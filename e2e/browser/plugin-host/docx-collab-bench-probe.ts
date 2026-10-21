export interface DocxCollabBenchResult {
  frames: number;
  ms: number;
  paints: number;
  /** Main-thread time inside the transport's event dispatch, summed over the run. */
  dispatchMs: number;
}

export interface DocxCollabBenchProbe {
  errors: string[];
  ready: Promise<void>;
  /** Opens the editor, replays `edits` peer edits as one frame per task and answers its SyncStep1 last. */
  join(edits: number): Promise<DocxCollabBenchResult>;
  /** Delivers `edits` peer edits to the synced editor, one frame per task, every `paceMs` (0: back to back). */
  type(edits: number, paceMs: number): Promise<DocxCollabBenchResult>;
}

declare global {
  interface Window {
    __docxCollabBench: DocxCollabBenchProbe;
  }
}
