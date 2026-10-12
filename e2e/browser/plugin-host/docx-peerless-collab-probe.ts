export interface DocxPeerlessCollaborationProbe {
  errors: string[];
  ready: Promise<void>;
  text(index: number): Promise<string>;
  painted(index: number): string;
  generation(index: number): number;
  synced(index: number): boolean;
  hold(index: number): void;
  release(index: number): void;
  drop(index: number): void;
  disconnect(index: number): void;
  reconnect(index: number): void;
  room(): Promise<void>;
  swap(): Promise<void>;
}

declare global { interface Window { __docxPeerlessCollab: DocxPeerlessCollaborationProbe } }
