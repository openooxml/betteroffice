import type { SessionScope } from '../../../../shared/office-session/transport';
import { bootWorker, type WorkerBootScope } from '../../../../shared/office-session/workerBoot';

bootWorker(
  self as unknown as SessionScope & WorkerBootScope,
  () => import('./workerStart').then((module) => module.startSessionWorker),
  (message) => ({ protocol: 1, kind: 'failure', code: 'crash', message }),
  'betteroffice.xlsx.worker',
);
