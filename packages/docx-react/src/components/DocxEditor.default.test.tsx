import { afterEach, beforeEach, expect, test } from 'bun:test';
import { chooseDocxEditorRoute, resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './DocxEditor/internals/engineChoice';
import { LEGACY_SYNC_API_PROP } from './DocxEditor/peerless/contracts';

beforeEach(() => setMissingWorkerCapabilitiesForTests([]));
afterEach(resetEngineChoiceForTests);

test('the default route requires no sync opt-in', () => {
  expect(chooseDocxEditorRoute({})).toEqual({ route: 'peerless' });
  expect(chooseDocxEditorRoute({ experimentalWorkerOpen: true })).toEqual({ route: 'peerless' });
});

test('explicit runtime choices precede collaboration, capabilities and sync callbacks', () => {
  setMissingWorkerCapabilitiesForTests(['Worker']);
  const props = { collaboration: {}, onChange() {}, experimentalWorkerOpen: false as const };
  expect(chooseDocxEditorRoute({ ...props, [LEGACY_SYNC_API_PROP]: true })).toMatchObject({ route: 'hybrid', warning: expect.stringContaining('0.6.0') });
  expect(chooseDocxEditorRoute(props)).toEqual({ route: 'in-thread' });
  expect(chooseDocxEditorRoute({ collaboration: {}, onChange() {} })).toEqual({ route: 'capability-fallback' });
  expect(chooseDocxEditorRoute({ onChange() {} })).toEqual({ route: 'capability-fallback' });
});

test('sync callbacks are refused only on the default worker route', () => {
  expect(chooseDocxEditorRoute({ onChange() {} })).toEqual({ route: 'sync-api-error' });
  expect(chooseDocxEditorRoute({ onChange: undefined })).toEqual({ route: 'peerless' });
});

test('collaboration takes the default worker route', () => {
  expect(chooseDocxEditorRoute({ collaboration: {} })).toEqual({ route: 'peerless' });
  expect(chooseDocxEditorRoute({ collaboration: { initialUpdate: new Uint8Array([0]) } })).toEqual({ route: 'peerless' });
  expect(chooseDocxEditorRoute({ collaboration: {}, onChange() {} })).toEqual({ route: 'sync-api-error' });
});
