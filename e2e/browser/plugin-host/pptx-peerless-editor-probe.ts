import type { PptxReadResult, PresentationCommandSelected } from '../../../packages/pptx/src/index';

export interface PeerlessEditorProbe {
  ready: Promise<void>;
  errors: string[];
  saves: number[][];
  setReadOnly(value: boolean, toolbar?: boolean): void;
  mainThreadEngines(): number;
  generation(editor?: number): number;
  hydrated(editor?: number): boolean;
  content(editor?: number): Promise<Extract<PptxReadResult, { ok: true }>>;
  selection(editor?: number): PresentationCommandSelected | null;
  activeSlide(editor?: number): number;
  paintedTexts(editor?: number): string[];
  point(position: number, editor?: number): { x: number; y: number };
  shapePoint(editor?: number): { x: number; y: number; units: { x: number; y: number } };
  select(start: number, end: number, editor?: number): Promise<boolean>;
  flush(editor?: number): Promise<void>;
  hold(editor?: number): void;
  release(editor?: number): void;
  held(editor?: number): number;
  saveAndReopen(editor?: number): Promise<number>;
  merge(): Promise<void>;
  queueRemoteMerge(): Promise<void>;
  settleRemoteMerge(): Promise<void>;
}

declare global {
  var __pptxMainEngineProbe: { capture(): void; count(): number } | undefined;
  interface Window { __pptxPeerlessEditor: PeerlessEditorProbe }
}
