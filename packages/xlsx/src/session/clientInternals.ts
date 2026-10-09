import type { Promisified } from '../../../../shared/office-session/types';
import type { OpenWorkbookOptions, WorkbookCalculationContext, WorkbookHandle } from '../wasm/loader';
import type { WorkbookSession } from './client';
import type { WorkbookInternalSessionMethods } from './replay';

export interface WorkbookPeerSource {
  bytes?: Uint8Array<ArrayBuffer>;
  options: OpenWorkbookOptions;
  module?: WebAssembly.Module;
  hydration?: string;
  initialCalculation?: WorkbookCalculationContext | null;
  version?: string;
  sequence?: number;
  disposed: boolean;
  pending?: Promise<WorkbookHandle>;
  snapshot?: Pick<Promisified<WorkbookInternalSessionMethods>,
    'beginPeerSnapshot' | 'pullPeerSnapshot' | 'endPeerSnapshot'>;
}

export const workbookPeerSources = new WeakMap<WorkbookSession, WorkbookPeerSource>();
