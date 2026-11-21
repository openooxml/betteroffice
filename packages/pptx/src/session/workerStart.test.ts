import { expect, test } from 'bun:test';
import { FakeWorkerScope } from '../../../../shared/office-session/testing/workerScope';
import { startSessionWorker } from './workerStart';

test('the presentation worker listens for messages before start returns, once per scope', () => {
  const scope = new FakeWorkerScope();
  startSessionWorker(scope);
  const listening = scope.listened.filter((type) => type === 'message').length;
  expect(listening).toBeGreaterThanOrEqual(2);
  startSessionWorker(scope);
  expect(scope.listened.filter((type) => type === 'message')).toHaveLength(listening);
  expect(scope.posted).toEqual([]);
});
