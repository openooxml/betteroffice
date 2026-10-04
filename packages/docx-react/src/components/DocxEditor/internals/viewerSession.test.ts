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

function session(workerProposals: boolean, generation = 1, workerOpen = true, mediaTokens = false) {
  return renderHook(
    (props: { workerOpen: boolean; workerProposals: boolean; generation: number; mediaTokens?: boolean }) =>
      useViewerSession(props.workerOpen && !props.mediaTokens, props.workerProposals, props.generation),
    { initialProps: { workerOpen, workerProposals, generation, mediaTokens } }
  );
}

test('a document opened read-only in the worker is a viewer session', () => {
  expect(session(true).result.current).toBe(true);
  expect(session(true, 1, false).result.current).toBe(false);
  expect(session(false).result.current).toBe(false);
});

test('a session opened for editing keeps the editing input when switched to viewing', () => {
  const { result, rerender } = session(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1, mediaTokens: false });
  expect(result.current).toBe(false);
});

test('a viewer session switched to editing keeps the editing input when switched back', () => {
  const { result, rerender } = session(true);
  rerender({ workerOpen: true, workerProposals: false, generation: 1, mediaTokens: false });
  expect(result.current).toBe(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1, mediaTokens: false });
  expect(result.current).toBe(false);
});

test('a viewer session leaves the viewer input when worker-open is turned off', () => {
  const { result, rerender } = session(true);
  rerender({ workerOpen: false, workerProposals: true, generation: 1, mediaTokens: false });
  expect(result.current).toBe(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1, mediaTokens: false });
  expect(result.current).toBe(false);
});

test('the next document opened while viewing is a viewer session again', () => {
  const { result, rerender } = session(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1, mediaTokens: false });
  rerender({ workerOpen: true, workerProposals: true, generation: 2, mediaTokens: false });
  expect(result.current).toBe(true);
});

test('a media-token document stays editor kind for its generation', () => {
  const { result, rerender } = session(true, 1, true, true);
  expect(result.current).toBe(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1, mediaTokens: false });
  expect(result.current).toBe(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 2, mediaTokens: false });
  expect(result.current).toBe(true);
});

test('enabling media tokens leaves viewer kind permanently for the current document', () => {
  const { result, rerender } = session(true);
  rerender({ workerOpen: true, workerProposals: true, generation: 1, mediaTokens: true });
  expect(result.current).toBe(false);
  rerender({ workerOpen: true, workerProposals: true, generation: 1, mediaTokens: false });
  expect(result.current).toBe(false);
});
