import { preloadResidentEngineWorker } from './residentEngineWorkerClient';

/**
 * Preload one spare resident editing worker. Opt-in.
 * @experimental
 */
export async function preloadDocxEngine(): Promise<void> {
  await preloadResidentEngineWorker();
}
