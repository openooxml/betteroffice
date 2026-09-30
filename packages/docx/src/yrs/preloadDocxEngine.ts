import { preloadResidentEngineWorker } from './residentEngineWorkerClient';

/**
 * Preload the editing engine and one spare resident worker. Opt-in.
 * @experimental
 */
export async function preloadDocxEngine(): Promise<void> {
  const main = [
    import('./wasm/index').then((wasm) => wasm.preloadEditWasm()),
    import('../docx/wasm').then((opc) => opc.preloadOpcWasm()),
    import('../docx/parseWasm').then((parse) => parse.preloadParseWasm()),
    import('../layout/wasm/index').then((layout) => layout.preloadLayoutWasm()),
  ];
  let worker: Promise<void>;
  try {
    worker = preloadResidentEngineWorker();
  } catch (error) {
    worker = Promise.reject(error);
  }
  await Promise.all([...main, worker]);
}
