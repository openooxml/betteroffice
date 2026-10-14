import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { WorkerDocumentReader } from './WorkerDocumentReader';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

function fixture() {
  const identities = mock(async () => ({ paragraphs: Array.from({ length: 70 }, (_, index) => ({
    session: { story: index < 65 ? 'body' : 'cell', paraId: String(index) },
  })) }));
  const read = mock(async (request: { story?: string; paraIds?: readonly string[] }) => ({ ok: true, version: 'native',
    paragraphs: (request.paraIds ?? []).map((paraId) => ({ story: request.story, paraId, text: `Unbuilt paragraph ${paraId}` })),
  }));
  return { identities, read, controller: { ref: { getParagraphIdentities: identities, readParagraphs: read } } as never };
}

test('accessible reader reaches never-built pages and child stories through bounded worker reads', async () => {
  const { controller, read } = fixture();
  const view = render(<WorkerDocumentReader controller={controller} version={0} />);
  expect(read).not.toHaveBeenCalled();
  fireEvent.click(view.getByRole('button', { name: 'Read document text' }));
  await waitFor(() => expect(view.getByText('Unbuilt paragraph 31')).toBeTruthy());
  expect(read.mock.calls.every(([request]) => request.paraIds!.length <= 32)).toBe(true);
  fireEvent.click(view.getByRole('button', { name: 'Next text' }));
  await waitFor(() => expect(view.getByText('Unbuilt paragraph 63')).toBeTruthy());
  fireEvent.click(view.getByRole('button', { name: 'Next text' }));
  await waitFor(() => expect(view.getByText('Unbuilt paragraph 69')).toBeTruthy());
  expect(read.mock.calls.at(-1)![0].story).toBe('cell');
  expect(view.getByRole('button', { name: 'Next text' }).hasAttribute('disabled')).toBe(true);
});

test('reader clears old text when document version changes', async () => {
  const { controller, read } = fixture();
  const view = render(<WorkerDocumentReader controller={controller} version={0} />);
  fireEvent.click(view.getByRole('button', { name: 'Read document text' }));
  await waitFor(() => expect(view.getByText('Unbuilt paragraph 0')).toBeTruthy());
  read.mockImplementation(async () => ({ ok: true, version: 'new', paragraphs: [] }));
  view.rerender(<WorkerDocumentReader controller={controller} version={1} />);
  await waitFor(() => expect(view.queryByText('Unbuilt paragraph 0')).toBeNull());
});
