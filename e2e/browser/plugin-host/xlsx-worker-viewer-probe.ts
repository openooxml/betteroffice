import type { CanvasPixelComparison } from './xlsx-canvas-parity';

export type ViewerArm = 'in-thread' | 'worker';

export interface ViewerPixelComparison extends CanvasPixelComparison {
  sheet: number;
  zoom: number;
  dpr: number;
  width: number;
  height: number;
  scrollLeft: number;
  scrollTop: number;
}

export interface WorkerViewerContract {
  handleIsNull: boolean;
  syncSaveIsNull: boolean;
  syncSelection: boolean;
  asyncSelection: boolean;
  scrollBefore: { left: number; top: number };
  scrollAfter: { left: number; top: number };
  targetPaintedAtResolution: boolean;
  targetVisibleAtResolution: boolean;
  savedLength: number;
  signature: number[];
}

export interface WorkerViewerProbe {
  ready: Promise<{ sheetCount: number }>;
  errors: string[];
  show(zoom: number, position: 'origin' | 'scrolled'): Promise<ViewerPixelComparison>;
  compareCurrent(sheet: number): Promise<ViewerPixelComparison>;
  contract(): Promise<WorkerViewerContract>;
}

declare global {
  interface Window { __xlsxWorkerViewer: WorkerViewerProbe }
}
