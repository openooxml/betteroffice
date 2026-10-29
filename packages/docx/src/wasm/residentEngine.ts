import type { WasmModuleState } from './loadWasmAsset';

interface ResidentEngine {
  state: WasmModuleState;
  glue: object;
}

let engine: ResidentEngine | undefined;

export function registerResidentEngine(state: WasmModuleState, glue: object): void {
  engine = { state, glue };
}

export function residentEngine(): ResidentEngine | undefined {
  return engine;
}
