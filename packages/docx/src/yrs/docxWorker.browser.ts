import { bootWorker } from '../../../../shared/office-session/workerBoot';

bootWorker(
  self as unknown as DedicatedWorkerGlobalScope,
  () => import('./residentEngineWorker').then((module) => module.startWorker),
  (error) => ({ id: 0, ok: false, terminal: true, failureCode: 'crash', error }),
  'betteroffice.docx.worker',
);
