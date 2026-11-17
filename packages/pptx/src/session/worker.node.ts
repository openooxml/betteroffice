import { sessionWorkerScope } from '../../../../shared/office-session/nodeSessionScope';
import { startSessionWorker } from './workerStart';

startSessionWorker(sessionWorkerScope());
