export interface DocxRelayJoinResult {
  ms: number;
  paints: number;
  merges: number;
  replayed: number;
  text: string;
}

export interface DocxRelayJoinProbe {
  errors: string[];
  ready: Promise<void>;
  retained(): number;
  text(index: number): Promise<string>;
  join(legacy: boolean, expected: string): Promise<DocxRelayJoinResult>;
}

declare global {
  interface Window {
    __docxRelayJoin: DocxRelayJoinProbe;
  }
}
