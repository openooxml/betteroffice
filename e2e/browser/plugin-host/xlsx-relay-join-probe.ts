export interface XlsxRelayJoinResult {
  ms: number;
  paints: number;
  clears: number;
  applies: number;
  replayed: number;
  longTaskMs: number;
  maxFrameGapMs: number;
  slowFrames: number;
}

export interface XlsxRelayJoinProbe {
  errors: string[];
  ready: Promise<void>;
  retained(): number;
  painted(index: number): string[];
  join(legacy: boolean, expected: string[]): Promise<XlsxRelayJoinResult>;
  watch(expected: string[]): Promise<XlsxRelayJoinResult>;
  mark(): void;
}

declare global {
  interface Window {
    __xlsxRelayJoin: XlsxRelayJoinProbe;
  }
}
