import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DocxProposalSnapshot, YrsSession } from '@betteroffice/docx/yrs';
import { yrsIdToNumericId } from '@betteroffice/docx/layout/render';
import { useHostProposalRevisions } from './useHostProposalRevisions';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function proposalSession() {
  let snapshot: DocxProposalSnapshot = { version: 'v1', previewVersion: 0, proposals: [] };
  const listeners = new Set<(snapshot: DocxProposalSnapshot) => void>();
  const session = {
    getProposals: () => snapshot,
    onProposalChange(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  } satisfies Pick<YrsSession, 'getProposals' | 'onProposalChange'>;
  return {
    session: session as YrsSession,
    listeners,
    publish(revisionIds: readonly string[], previewVersion = 0) {
      snapshot = {
        version: 'v1',
        previewVersion,
        proposals:
          revisionIds.length === 0
            ? []
            : [
                {
                  id: 'p1',
                  state: previewVersion === 0 ? 'proposed' : 'accepted',
                  paragraph: { kind: 'session', sessionId: 's1', story: 'body', paraId: 'p1' },
                  revisionIds,
                  changed: true,
                },
              ],
      };
      for (const listener of listeners) listener(snapshot);
    },
  };
}

test('proposal keys update on notifications, stay stable for decisions and follow the current session', () => {
  const first = proposalSession();
  const second = proposalSession();
  const view = renderHook(
    ({ session }: { session: YrsSession | null }) => useHostProposalRevisions(session),
    { initialProps: { session: first.session as YrsSession | null } }
  );
  expect(view.result.current.size).toBe(0);
  expect(first.listeners.size).toBe(1);

  act(() => first.publish(['host-a', '7', '7']));
  expect(view.result.current).toEqual(
    new Set([`revision-${yrsIdToNumericId('host-a')}`, 'revision-7'])
  );
  const keys = view.result.current;
  act(() => first.publish(['7', 'host-a'], 1));
  expect(view.result.current).toBe(keys);

  second.publish(['host-b']);
  view.rerender({ session: second.session });
  expect(view.result.current).toEqual(new Set([`revision-${yrsIdToNumericId('host-b')}`]));
  expect(first.listeners.size).toBe(0);
  expect(second.listeners.size).toBe(1);
  const secondKeys = view.result.current;
  act(() => first.publish(['old-session']));
  expect(view.result.current).toBe(secondKeys);

  act(() => second.publish([]));
  expect(view.result.current.size).toBe(0);
  const empty = view.result.current;
  act(() => second.publish([], 1));
  expect(view.result.current).toBe(empty);
  view.rerender({ session: null });
  expect(view.result.current).toBe(empty);
  expect(second.listeners.size).toBe(0);
});
