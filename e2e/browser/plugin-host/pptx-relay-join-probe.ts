export interface PptxRelayJoinResult {
  ms: number;
  paints: number;
  commits: number;
  applies: number;
  replayed: number;
  longTaskMs: number;
  maxFrameGapMs: number;
  slowFrames: number;
}

export interface PptxRelayJoinProbe {
  errors: string[];
  ready: Promise<void>;
  retained(): number;
  edit(count: number): Promise<string>;
  join(legacy: boolean, expected: string): Promise<PptxRelayJoinResult>;
}

declare global {
  interface Window {
    __pptxRelayJoin: PptxRelayJoinProbe;
  }
}
