import { expect, test } from 'bun:test';
import * as xlsx from './index';
import * as worker from './worker';

test('root re-exports the worker entry', () => {
  expect(Object.keys(worker).sort()).toEqual(['WorkbookPngUnavailableError', 'WorkbookRecoveryReplayError',
    'createWorkbookEditSession', 'openWorkbookSession', 'workbookSessionRecovery']);
  for (const [name, value] of Object.entries(worker)) {
    expect(xlsx[name as keyof typeof xlsx]).toBe(value);
  }
});
