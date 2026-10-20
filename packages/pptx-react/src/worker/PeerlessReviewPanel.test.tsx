import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import type { Proposal, ProposalChange } from '@betteroffice/pptx';
import type { PptxCommandStore } from '../commands/types';
import { PeerlessReview, type PeerlessReviewState } from './PeerlessReview';
import type { PeerlessPresentation } from './usePeerlessPresentation';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor, within } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function proposal(overrides: Partial<Proposal> = {}): Proposal {
  return {
    id: 'proposal', agentId: 'reviewer', note: 'Update the slide', edits: [], staleTargets: [],
    changes: [{ slideId: 'slide', shapeId: 'shape', before: null, after: null, oldText: 'Original', newText: 'Updated' }],
    ...overrides,
  };
}

function mount(proposals: Proposal[]) {
  const execute = mock(async () => ({ ok: true as const, status: 'executed' as const }));
  const commands = { execute } as unknown as PptxCommandStore;
  const run = {
    current: true, require() {}, getVisualSnapshot: () => null, subscribeVisual: () => () => {},
    client: { state: { slides: [{ id: 'slide' }], activeSlideId: 'slide' } },
    store: { getSnapshot: () => ({ composition: null }), fence: () => new Promise<void>(() => {}) },
  } as unknown as PeerlessPresentation;
  const onError = mock(() => {});
  const state: PeerlessReviewState = { open: true, selectedId: null, diff: false, presenting: false };
  const props = { run, state, proposals, commands, onError, exit() {} };
  const view = render(<PeerlessReview {...props} />);
  return { view, execute, props };
}

test('review heading counts proposals and updates when one is removed', async () => {
  const { view, props } = mount([proposal(), proposal({ id: 'second' })]);
  const panel = await view.findByTestId('pptx-proposals-panel');
  expect(panel.querySelector('strong')?.textContent).toBe('Agent proposals 2');
  view.rerender(<PeerlessReview {...props} proposals={[proposal()]} />);
  await waitFor(() => expect(panel.querySelector('strong')?.textContent).toBe('Agent proposals 1'));
});

test('review shows the author alongside the note and every text change', async () => {
  const first = proposal();
  const changes: ProposalChange[] = [...first.changes, { ...first.changes[0]!, shapeId: 'other', oldText: 'Second', newText: 'Replacement' }];
  const { view } = mount([proposal({ changes })]);
  await waitFor(() => expect(view.getByTestId('pptx-proposal').querySelector('strong')?.textContent).toBe('reviewer'));
  const card = view.getByTestId('pptx-proposal');
  expect(card.getAttribute('data-proposal-id')).toBe('proposal');
  expect(card.textContent).toContain('Update the slide');
  await waitFor(() => expect(Array.from(card.querySelectorAll('del'), (node) => node.textContent)).toEqual(['Original', 'Second']));
  expect(Array.from(card.querySelectorAll('ins'), (node) => node.textContent)).toEqual(['Updated', 'Replacement']);
});

test('review previews speaker notes with the empty text placeholder', async () => {
  const notes = { ...proposal().changes[0]!, shapeId: null, oldText: '', newText: 'Speaker notes' };
  const { view } = mount([proposal({ changes: [notes] })]);
  await waitFor(() => expect(view.getByTestId('pptx-proposal-notes-diff').querySelector('ins')?.textContent).toBe('Speaker notes'));
  expect(view.getByTestId('pptx-proposal-notes-diff').querySelector('del')?.textContent).toBe('Empty');
});

test('review warns about stale targets and disables ordinary acceptance', async () => {
  const { view } = mount([proposal({ staleTargets: ['shape'] })]);
  await waitFor(() => expect(view.getByTestId('pptx-proposal-stale').getAttribute('role')).toBe('alert'));
  expect(view.getByTestId('pptx-proposal-stale').textContent).toContain('changed after the proposal was created');
  expect((view.getByRole('button', { name: 'Accept proposal' }) as HTMLButtonElement).disabled).toBe(true);
  expect((view.getByRole('button', { name: 'Reject' }) as HTMLButtonElement).disabled).toBe(false);
  expect(view.queryByTestId('pptx-proposal-force')).toBeNull();
});

test('review withholds force acceptance until the fresh preview is ready', async () => {
  const { view, execute, props } = mount([proposal({ staleTargets: ['shape'] })]);
  fireEvent.click(await view.findByRole('button', { name: 'Preview changes' }));
  await waitFor(() => expect(execute).toHaveBeenCalledWith('proposalDiff', { enabled: true }));
  expect(execute).toHaveBeenCalledWith('proposalSelect', { proposalId: 'proposal' });
  view.rerender(<PeerlessReview {...props} state={{ ...props.state, selectedId: 'proposal', diff: true }} />);
  expect(view.queryByTestId('pptx-proposal-force') === null).toBe(true);
  expect(execute).not.toHaveBeenCalledWith('proposalAccept', { proposalId: 'proposal', force: true });
  expect(props.onError).not.toHaveBeenCalled();
});

test('review accepts and rejects each proposal independently', async () => {
  const { view, execute } = mount([proposal(), proposal({ id: 'second', agentId: 'other', note: null })]);
  await waitFor(() => expect(view.getAllByTestId('pptx-proposal')).toHaveLength(2));
  const [first, second] = view.getAllByTestId('pptx-proposal');
  await act(async () => {
    fireEvent.click(within(first!).getByRole('button', { name: 'Accept proposal' }));
    fireEvent.click(within(second!).getByRole('button', { name: 'Reject' }));
  });
  expect(execute).toHaveBeenCalledWith('proposalAccept', { proposalId: 'proposal' });
  expect(execute).toHaveBeenCalledWith('proposalReject', { proposalId: 'second' });
});
