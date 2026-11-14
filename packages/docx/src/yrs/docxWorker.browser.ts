import { startWorker } from './residentEngineWorker';

startWorker(self as unknown as DedicatedWorkerGlobalScope);
