import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { UnifiedSidebar } from '../components/UnifiedSidebar';
import type { TrackedChangeEntry } from '../components/sidebar/cardUtils';
import { useCommentSidebarItems, type CommentCallbacks } from './useCommentSidebarItems';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>(resolve => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const change: TrackedChangeEntry = { revisionId: 7, type: 'insertion', author: 'Reviewer', text: 'Added', from: 1, to: 6 };

test('memoized cards use current accept and reject handlers after an edit', () => {
  let items: ReturnType<typeof useCommentSidebarItems> = [];
  function Harness({ changes, callbacks }: { changes: TrackedChangeEntry[]; callbacks: CommentCallbacks }) {
    items = useCommentSidebarItems({ comments: [], trackedChanges: changes, callbacks });
    return <UnifiedSidebar items={items} anchorPositions={new Map()} renderedDomContext={null}
      pageWidth={600} zoom={1} editorContainerRef={{ current: null }} />;
  }
  const oldAccept = mock((_id: number) => {}), oldReject = mock((_id: number) => {});
  const accept = mock((_id: number) => {}), reject = mock((_id: number) => {});
  const oldCallbacks = { onAcceptChangeById: oldAccept, onRejectChangeById: oldReject };
  const view = render(<Harness changes={[change]} callbacks={oldCallbacks} />);
  const first = items[0];
  view.rerender(<Harness changes={[{ ...change }]} callbacks={{ ...oldCallbacks }} />);
  expect(items[0]).toBe(first);
  view.rerender(<Harness changes={[{ ...change, text: 'Edited', from: 2, to: 8 }]}
    callbacks={{ onAcceptChangeById: accept, onRejectChangeById: reject }} />);
  fireEvent.click(view.getByTitle('Accept change'));
  fireEvent.click(view.getByTitle('Reject change'));
  expect(accept).toHaveBeenCalledWith(7);
  expect(reject).toHaveBeenCalledWith(7);
  expect(oldAccept).not.toHaveBeenCalled();
  expect(oldReject).not.toHaveBeenCalled();
  const latestAccept = mock((_id: number) => {}), latestReject = mock((_id: number) => {});
  view.rerender(<Harness changes={[{ ...change, text: 'Edited', from: 2, to: 8 }]}
    callbacks={{ onAcceptChangeById: latestAccept, onRejectChangeById: latestReject }} />);
  fireEvent.click(view.getByTitle('Accept change'));
  fireEvent.click(view.getByTitle('Reject change'));
  expect(latestAccept).toHaveBeenCalledWith(7);
  expect(latestReject).toHaveBeenCalledWith(7);
  expect(accept).toHaveBeenCalledTimes(1);
  expect(reject).toHaveBeenCalledTimes(1);
});
