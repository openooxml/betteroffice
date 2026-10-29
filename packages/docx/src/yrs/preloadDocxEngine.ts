import { preloadResidentEngineWorker } from './residentEngineWorkerClient';

/**
 * Preload the editing engine and one spare resident worker. Opt-in.
 * @experimental
 */
export async function preloadDocxEngine(): Promise<void> {
  const main = import('./wasm/index').then((wasm) => wasm.preloadEditWasm());
  let worker: Promise<void>;
  try {
    worker = preloadResidentEngineWorker();
  } catch (error) {
    worker = Promise.reject(error);
  }
  await Promise.all([main, worker]);
}
