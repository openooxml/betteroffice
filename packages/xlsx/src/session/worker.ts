import type { SessionScope } from '../../../../shared/office-session/transport';
import { startSessionWorker } from './workerStart';

startSessionWorker(self as unknown as SessionScope);
