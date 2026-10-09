export type ViewerArm = 'in-thread' | 'worker';

export interface PixelComparison {
  slide: number;
  zoom: number;
  dpr: number;
  thumbnail: boolean;
  differingPixels: number;
  maxChannelDelta: number;
  localImageDraws: number;
  workerImageDraws: number;
  localImagePixels: number;
  workerImagePixels: number;
}

export interface WorkerViewerProbe {
  ready: Promise<{ slideCount: number; initialSlide: number }>;
  errors: string[];
  show(slide: number, zoom: number): Promise<PixelComparison>;
  compareCurrent(slide: number): Promise<PixelComparison>;
  thumbnails(): Promise<PixelComparison[]>;
  workerImage(): Promise<{ draws: number; pixels: number }>;
}

declare global {
  interface Window { __pptxWorkerViewer: WorkerViewerProbe }
}
