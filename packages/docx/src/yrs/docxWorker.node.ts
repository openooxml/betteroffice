import { sessionWorkerScope } from '../../../../shared/office-session/nodeSessionScope';
import { startWorker } from './residentEngineWorker';

startWorker(sessionWorkerScope() as unknown as DedicatedWorkerGlobalScope);
