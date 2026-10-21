import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useRef, useState } from 'react';
import { resolvePresenceColor, type CollaborationPeer } from '@betteroffice/docx/collaboration';
import { TrackedChangeCard } from './TrackedChangeCard';
import { ReviewSummary, TrackedChangesReview } from './TrackedChangesReview';
import { decideReviewChange, resolveReviewAuthor, reviewRevisionIds, TrackedChangesContext, type TrackedChangesOptions } from './trackedChanges';
import type { TrackedChangeEntry } from './cardUtils';
import { useCommentSidebarItems } from '../../hooks/useCommentSidebarItems';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>(resolve => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const change: TrackedChangeEntry = { revisionId: 7, type: 'replacement', author: 'Alex Reed', date: '2026-10-01T09:00:00Z',
  text: 'four weeks', deletedText: 'six weeks', insertionRevisionId: 8, coalescedRevisionIds: [7, 9], from: 10, to: 20 };
const sidebarProps = { isExpanded: false, onToggleExpand() {}, measureRef() {} };
const peer = { clientId: 3, user: { name: change.author, color: '#123456' } } as CollaborationPeer;

test('authors use peer, local and stable palette colours with host overrides and default icons', () => {
  expect(resolveReviewAuthor(change.author, 'human', undefined, [peer]).color).toBe('#123456');
  expect(resolveReviewAuthor(change.author, 'human', undefined, [peer], { name: change.author, color: '#abcdef' }).color).toBe('#123456');
  expect(resolveReviewAuthor(change.author, 'human', undefined, [], { name: change.author, color: '#abcdef' }).color).toBe('#abcdef');
  expect(resolveReviewAuthor(change.author, 'human', undefined, [], { name: change.author }, 3).color).toBe(resolvePresenceColor(3, undefined));
  const fallback = resolveReviewAuthor(change.author);
  let hash = 0;
  for (const char of change.author) hash = (Math.imul(hash, 31) + char.charCodeAt(0)) >>> 0;
  expect(fallback.color).toBe(resolvePresenceColor(hash, undefined));
  expect(fallback.icon).toBe('AR');
  const getAuthorAppearance = mock(() => ({ color: '#fedcba', icon: <b>R</b>, kind: 'agent' as const }));
  const custom = resolveReviewAuthor(change.author, 'human', { getAuthorAppearance }, [peer]);
  expect(getAuthorAppearance).toHaveBeenCalledWith({ name: change.author, kind: 'human' });
  expect(custom).toMatchObject({ color: '#fedcba', kind: 'agent' });
  const view = render(<><span>{custom.icon}</span><span>{resolveReviewAuthor('Assistant', 'agent').icon}</span></>);
  expect(view.container.querySelector('b')?.textContent).toBe('R');
  expect(view.container.querySelector('.docx-review-spark path')?.getAttribute('d')).toContain('M8 .8');
});

test('cards expose their parts, strike and insertion diff, always-visible by-id actions and active navigation', () => {
  const accept = mock((_id: number) => {}), reject = mock((_id: number) => {}), next = mock(() => {}), previous = mock(() => {});
  const range = mock(() => {});
  const view = render(<TrackedChangeCard {...sidebarProps} change={change} replies={[]} onAcceptById={accept} onRejectById={reject} onAccept={range}
    onNext={next} onPrevious={previous} index={1} total={3} />);
  const root = view.container.querySelector('.docx-tracked-change-card.docx-review-card')!;
  expect(root.getAttribute('data-active')).toBe('false');
  expect(root.getAttribute('data-kind')).toBe('human');
  expect(root.getAttribute('data-revision-id')).toBe('7');
  for (const part of ['header', 'name', 'meta', 'diff', 'actions']) expect(root.querySelector(`.docx-review-card__${part}`)).not.toBeNull();
  expect(root.querySelector('del.docx-review-diff-del')?.textContent).toBe('six weeks');
  expect(root.querySelector('ins.docx-review-diff-ins')?.textContent).toBe('four weeks');
  expect(root.querySelector('.docx-review-card__diff')?.textContent).toBe('six weeks → four weeks');
  fireEvent.click(view.getByRole('button', { name: 'Accept', exact: true }));
  fireEvent.click(view.getByRole('button', { name: 'Reject', exact: true }));
  expect(accept.mock.calls.map(call => call[0])).toEqual([7, 8, 9]);
  expect(reject.mock.calls.map(call => call[0])).toEqual([7, 8, 9]);
  expect(range).not.toHaveBeenCalled();
  expect(root.querySelector('.docx-review-nav')).toBeNull();
  view.rerender(<TrackedChangeCard {...sidebarProps} isExpanded change={change} replies={[]} onNext={next} onPrevious={previous} index={1} total={3} />);
  expect(view.container.querySelector('.docx-review-card')?.getAttribute('data-active')).toBe('true');
  expect(view.container.querySelector('.docx-review-nav__count')?.textContent).toBe('2/3');
  fireEvent.click(view.getByRole('button', { name: 'Previous change' }));
  fireEvent.click(view.getByRole('button', { name: 'Next change' }));
  expect(previous).toHaveBeenCalledTimes(1);
  expect(next).toHaveBeenCalledTimes(1);
});

test('insertions, deletions and structural revisions keep their labels and colours', () => {
  const view = render(<TrackedChangeCard {...sidebarProps} change={{ ...change, type: 'insertion' }} replies={[]} />);
  expect(view.container.querySelector('.docx-review-card__diff')?.textContent).toBe('Added four weeks');
  expect(view.container.querySelector('ins.docx-review-diff-ins')).not.toBeNull();
  view.rerender(<TrackedChangeCard {...sidebarProps} change={{ ...change, type: 'deletion' }} replies={[]} />);
  expect(view.container.querySelector('.docx-review-card__diff')?.textContent).toBe('Deleted four weeks');
  expect(view.container.querySelector('del.docx-review-diff-del')).not.toBeNull();
  view.rerender(<TrackedChangeCard {...sidebarProps} change={{ ...change, type: 'rowInserted' }} replies={[]} />);
  expect(view.container.querySelector('.docx-review-diff-ins')?.textContent).toBe('Inserted row');
});

test('appearance options reach the built-in card and full card override with a working measureRef', () => {
  const measure = mock((_element: HTMLDivElement | null) => {});
  const accept = mock((_id: number) => {});
  const options: TrackedChangesOptions = { getAuthorAppearance: () => ({ icon: <b data-testid="author-icon">A</b>, color: '#778899', kind: 'agent' }) };
  const view = render(<TrackedChangesContext.Provider value={{ options }}><TrackedChangeCard {...sidebarProps} change={change} replies={[]} /></TrackedChangesContext.Provider>);
  expect(view.getByTestId('author-icon')).not.toBeNull();
  expect((view.container.querySelector('.docx-review-card') as HTMLElement).style.getPropertyValue('--doc-review-author')).toBe('#778899');
  expect(view.container.querySelector('.docx-review-avatar')?.getAttribute('data-kind')).toBe('agent');
  const override = mock((props: Parameters<NonNullable<TrackedChangesOptions['renderCard']>>[0]) =>
    <div ref={props.measureRef} data-testid="custom-card"><button onClick={props.accept}>{props.author.name}</button></div>);
  view.rerender(<TrackedChangesContext.Provider value={{ options: { ...options, renderCard: override } }}>
    <TrackedChangeCard {...sidebarProps} measureRef={measure} change={change} replies={[]} onAcceptById={accept} />
  </TrackedChangesContext.Provider>);
  expect(view.container.querySelector('.docx-review-card')).toBeNull();
  expect(measure).toHaveBeenCalledWith(view.getByTestId('custom-card'));
  expect(override.mock.calls[0]![0].author.color).toBe('#778899');
  fireEvent.click(view.getByRole('button', { name: change.author }));
  expect(accept.mock.calls.map(call => call[0])).toEqual([7, 8, 9]);
});

test('summary orientations count changes, cap distinct avatars and accept every revision', () => {
  const accepted: number[] = [];
  const changes = [change, { ...change, revisionId: 10, insertionRevisionId: 11, coalescedRevisionIds: [] }];
  const authors = ['A', 'B', 'A', 'C', 'D'].map(name => resolveReviewAuthor(name));
  const acceptAll = () => changes.forEach(change => decideReviewChange(change, id => { accepted.push(id); }));
  const view = render(<ReviewSummary authors={authors} count={2} orientation="horizontal" acceptAll={acceptAll} />);
  expect(view.container.querySelector('.docx-review-summary')?.getAttribute('data-orientation')).toBe('horizontal');
  expect(view.container.querySelectorAll('.docx-review-avatar')).toHaveLength(3);
  expect(view.container.querySelector('.docx-review-summary__count')?.textContent).toBe('2 changes');
  expect(view.getByRole('button', { name: 'Accept all changes' }).textContent).toBe('Accept all');
  fireEvent.click(view.getByRole('button', { name: 'Accept all changes' }));
  expect(accepted).toEqual([7, 8, 9, 10, 11]);
  expect(reviewRevisionIds(change)).toEqual([7, 8, 9]);
  view.rerender(<ReviewSummary authors={authors} count={2} orientation="vertical" acceptAll={acceptAll} />);
  expect(view.container.querySelector('.docx-review-summary')?.getAttribute('data-orientation')).toBe('vertical');
  expect(view.container.querySelector('.docx-review-summary__count')?.textContent).toBe('2');
  expect(view.getByRole('button', { name: 'Accept all changes' }).querySelector('svg')).not.toBeNull();
});

function DotsHarness({ options = { display: 'dots' }, readOnly = false }: { options?: TrackedChangesOptions; readOnly?: boolean }) {
  const host = useRef<HTMLDivElement>(null), scroller = useRef<HTMLDivElement>(null);
  const [active, setActive] = useState<string | null>(null);
  const items = useCommentSidebarItems({ comments: [], trackedChanges: [change, { ...change, revisionId: 20, author: 'Riley', from: 30, to: 40 }], callbacks: {} });
  const rects = new Map(items.map(item => [item.id, [{ pageIndex: 0, x: 30, y: 100, width: 50, height: 12 }]]));
  return <div ref={scroller}><div ref={host}><TrackedChangesReview items={items} options={options} pages={[{ page: { pageIndex: 0 }, bounds: { x: 0, y: 0, width: 600, height: 800 } }]}
    rects={rects} zoom={1} pageWidth={600} sidebarOpen={false} hostRef={host} scrollerRef={scroller} activeId={active} setActive={setActive}
    select={() => {}} readOnly={readOnly} acceptAll={() => {}}>{options.display !== 'dots' && items.map(item => <div key={item.id}>{item.render({ ...sidebarProps, isExpanded: item.id === active, onToggleExpand: () => setActive(item.id) })}</div>)}</TrackedChangesReview></div></div>;
}

test('dots open on hover and focus, survive the balloon grace period, close on Escape and navigate', async () => {
  const view = render(<DotsHarness />);
  const dots = await waitFor(() => { const buttons = view.container.querySelectorAll<HTMLButtonElement>('.docx-review-dot'); expect(buttons).toHaveLength(2); return buttons; });
  expect(view.container.querySelector('.docx-review-card')).toBeNull();
  expect(dots[0]!.getAttribute('aria-label')).toBe('Change by Alex Reed');
  expect(dots[0]!.getAttribute('aria-expanded')).toBe('false');
  expect(Number.parseFloat(dots[1]!.style.left) - Number.parseFloat(dots[0]!.style.left)).toBe(22);
  fireEvent.mouseEnter(dots[0]!);
  expect(view.container.querySelector('.docx-review-balloon .docx-review-card')).not.toBeNull();
  expect(view.container.querySelector('svg.docx-review-leader')).not.toBeNull();
  fireEvent.mouseLeave(dots[0]!);
  fireEvent.mouseEnter(view.container.querySelector('.docx-review-balloon')!);
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 180)); });
  expect(view.container.querySelector('.docx-review-balloon')).not.toBeNull();
  fireEvent.mouseLeave(view.container.querySelector('.docx-review-balloon')!);
  await waitFor(() => expect(view.container.querySelector('.docx-review-balloon')).toBeNull());
  fireEvent.focus(dots[0]!);
  expect(view.container.querySelector('.docx-review-balloon')).not.toBeNull();
  fireEvent.keyDown(dots[0]!, { key: 'Escape' });
  expect(view.container.querySelector('.docx-review-balloon')).toBeNull();
  expect(document.activeElement).toBe(dots[0]);
  fireEvent.click(dots[0]!);
  fireEvent.mouseLeave(dots[0]!);
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 180)); });
  expect(view.container.querySelector('.docx-review-balloon')).not.toBeNull();
  fireEvent.click(view.getByRole('button', { name: 'Next change' }));
  expect(view.container.querySelector('.docx-review-nav__count')?.textContent).toBe('2/2');
  expect(dots[1]!.getAttribute('aria-expanded')).toBe('true');
});

test('dots balloon uses renderCard; cards have no dots; read-only hides the summary', async () => {
  const override = mock((props: Parameters<NonNullable<TrackedChangesOptions['renderCard']>>[0]) => <div ref={props.measureRef} data-testid="balloon-override">{props.index + 1}/{props.total}</div>);
  const view = render(<DotsHarness options={{ display: 'dots', renderCard: override }} />);
  fireEvent.focus(await view.findByRole('button', { name: 'Change by Alex Reed' }));
  expect(view.getByTestId('balloon-override').textContent).toBe('1/2');
  expect(override.mock.calls[0]![0].isActive).toBe(true);
  view.rerender(<DotsHarness options={{ display: 'cards' }} readOnly />);
  expect(view.container.querySelector('.docx-review-dot')).toBeNull();
  expect(view.container.querySelector('.docx-review-summary')).toBeNull();
  expect(view.container.querySelectorAll('.docx-review-card')).toHaveLength(2);
});

test('cards show the horizontal summary and an author-coloured leader for the active card', async () => {
  const view = render(<DotsHarness options={{ display: 'cards' }} />);
  await waitFor(() => expect(view.container.querySelector('.docx-review-summary')?.getAttribute('data-orientation')).toBe('horizontal'));
  expect(view.container.querySelector('.docx-review-dot')).toBeNull();
  fireEvent.click(view.container.querySelector('.docx-review-card')!);
  const leader = await waitFor(() => { const leader = view.container.querySelector<SVGElement>('svg.docx-review-leader'); expect(leader).not.toBeNull(); return leader!; });
  expect(leader.style.getPropertyValue('--doc-review-author')).toBe(resolveReviewAuthor(change.author).color);
  expect(leader.querySelector('path')?.getAttribute('d')).toContain(' L');
  expect(view.container.querySelector('.docx-review-nav__count')?.textContent).toBe('1/2');
});

test('balloons flip inside the visible right edge and the summary follows viewport scrolling', async () => {
  const view = render(<DotsHarness />);
  const scroller = view.container.firstElementChild as HTMLDivElement;
  const host = scroller.firstElementChild as HTMLDivElement;
  let right = 1600, top = 0;
  scroller.getBoundingClientRect = () => ({ left: 0, top: 0, right, bottom: 800, width: right, height: 800 }) as DOMRect;
  host.getBoundingClientRect = () => ({ left: 0, top, right: 1000, bottom: top + 1000, width: 1000, height: 1000 }) as DOMRect;
  Object.defineProperty(host, 'clientWidth', { value: 1000 });
  fireEvent.resize(window);
  fireEvent.focus(await view.findByRole('button', { name: 'Change by Alex Reed' }));
  const balloon = () => view.container.querySelector<HTMLElement>('.docx-review-balloon')!;
  const gutter = () => view.container.querySelector<HTMLElement>('.docx-review-gutter')!;
  expect(Number.parseFloat(balloon().style.left)).toBeGreaterThan(Number.parseFloat(gutter().style.left));
  right = 1000;
  fireEvent.resize(window);
  expect(Number.parseFloat(balloon().style.left)).toBeLessThan(Number.parseFloat(gutter().style.left));
  top = -200;
  fireEvent.scroll(scroller);
  expect(view.container.querySelector<HTMLElement>('.docx-review-summary-position')?.style.top).toBe('212px');
});
