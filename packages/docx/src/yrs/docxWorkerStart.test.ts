import { expect, test } from 'bun:test';
import { FakeWorkerScope } from '../../../../shared/office-session/testing/workerScope';
import { startWorker } from './residentEngineWorker';

test('the document worker installs its message handlers before start returns, once per scope', () => {
  const scope = new FakeWorkerScope();
  startWorker(scope as unknown as DedicatedWorkerGlobalScope);
  expect(scope.assigned).toEqual(['message', 'messageerror']);
  expect(scope.onmessage).toBeFunction();
  expect(scope.onmessageerror).toBeFunction();
  const handler = scope.onmessage;
  startWorker(scope as unknown as DedicatedWorkerGlobalScope);
  expect(scope.assigned).toEqual(['message', 'messageerror']);
  expect(scope.onmessage).toBe(handler);
});
