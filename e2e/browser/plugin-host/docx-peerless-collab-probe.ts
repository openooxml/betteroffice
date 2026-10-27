export interface DocxPeerlessCollaborationProbe {
  errors: string[];
  ready: Promise<void>;
  load(bytes: number[]): Promise<void>;
  text(index: number, view?: 'accepted' | 'original'): Promise<string>;
  painted(index: number): string;
  generation(index: number): number;
  clientId(index: number): number | null;
  stateVector(index: number): Promise<Array<[number, number]>>;
  synced(index: number): boolean;
  hold(index: number): void;
  release(index: number): void;
  drop(index: number): void;
  heldUpdates(index: number): number;
  disconnect(index: number): void;
  /** Releases held frames and disconnects as soon as one of their worker merges is dispatched. */
  dropMidMerge(index: number): Promise<void>;
  reconnect(index: number): void;
  restart(): void;
  checkpoint(index: number): Promise<void>;
  rejoin(index: number): Promise<void>;
  propose(index: number, text: string): Promise<void>;
  settle(index: number, state: 'accepted' | 'rejected'): Promise<void>;
  room(): Promise<void>;
  swap(): Promise<void>;
}

declare global { interface Window { __docxPeerlessCollab: DocxPeerlessCollaborationProbe } }
