export interface WorkerEditorProbe {
  ready: Promise<void>;
  errors: string[];
  previews: string[];
  previewFrames: string[];
  commitOrder: { kind: 'painted-preview' | 'mutator-entry'; text: string }[];
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
