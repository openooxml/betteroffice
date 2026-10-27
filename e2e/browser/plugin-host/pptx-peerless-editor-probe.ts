import type { PptxReadResult } from '../../../packages/pptx/src/index';
import type { PresentationCommandSelected } from '../../../packages/pptx/src/worker';
import type { Proposal, ProposalEdit } from '../../../packages/pptx/src/proposals';

export interface PeerlessEditorProbe {
  ready: Promise<void>;
  errors: string[];
  saves: number[][];
  setReadOnly(value: boolean, toolbar?: boolean): void;
  editor(editor?: number): {
    commands: { execute(id: string, args?: unknown): unknown; getState(id: string, args?: unknown): unknown };
    flushPendingInput(): Promise<void>;
    handleAsync(): Promise<{ propose(author: string, note: string, edits: ProposalEdit[]): Promise<Proposal>; listProposals(): Promise<Proposal[]> }>;
  };
  takeSaved(editor?: number): number[] | null;
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
