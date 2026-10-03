import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { useViewerSession } from './viewerSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function session(workerProposals: boolean, generation = 1, workerOpen = true) {
  return renderHook(
    (props: { workerOpen: boolean; workerProposals: boolean; generation: number }) =>
      useViewerSession(props.workerOpen, props.workerProposals, props.generation),
    { initialProps: { workerOpen, workerProposals, generation } }
  );
}

test('a document opened read-only in the worker is a viewer session', () => {
  expect(session(true).result.current).toBe(true);
  expect(session(true, 1, false).result.current).toBe(false);
  expect(session(false).result.current).toBe(false);
});

test('a session opened for editing keeps the editing input when switched to viewing', () => {
  const { result, rerender } = session(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1 });
  expect(result.current).toBe(false);
});

test('a viewer session switched to editing keeps the editing input when switched back', () => {
  const { result, rerender } = session(true);
  rerender({ workerOpen: true, workerProposals: false, generation: 1 });
  expect(result.current).toBe(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1 });
  expect(result.current).toBe(false);
});

test('a viewer session leaves the viewer input when worker-open is turned off', () => {
  const { result, rerender } = session(true);
  rerender({ workerOpen: false, workerProposals: true, generation: 1 });
  expect(result.current).toBe(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1 });
  expect(result.current).toBe(false);
});

test('the next document opened while viewing is a viewer session again', () => {
  const { result, rerender } = session(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1 });
  rerender({ workerOpen: true, workerProposals: true, generation: 2 });
  expect(result.current).toBe(true);
});
