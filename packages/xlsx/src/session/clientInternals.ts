import type { createWorkbookPeerOpener, OpenWorkbookOptions, WorkbookCalculationContext, WorkbookHandle } from '../wasm/loader';
import type { WorkbookSession } from './client';
import type { WorkbookPeerHydrationError } from './peerHydrationError';

export interface WorkbookPeerSource {
  bytes?: Uint8Array<ArrayBuffer>;
  options: OpenWorkbookOptions;
  module?: WebAssembly.Module;
  wasm?: ArrayBuffer | WebAssembly.Module;
  hydration: string[];
  receivedHydration: boolean;
  deliveryComplete: boolean;
  opener?: ReturnType<typeof createWorkbookPeerOpener>;
  wake?: () => void;
  holds: number;
  failure?: WorkbookPeerHydrationError;
  initialCalculation?: WorkbookCalculationContext | null;
  version?: string;
  sequence?: number;
  disposed: boolean;
  pending?: Promise<WorkbookHandle>;
}

export const workbookPeerSources = new WeakMap<WorkbookSession, WorkbookPeerSource>();
