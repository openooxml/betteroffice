import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useRef, useState } from 'react';
import { resolvePresenceColor, type CollaborationPeer } from '@betteroffice/docx/collaboration';
import { TrackedChangeCard } from './TrackedChangeCard';
import { ReviewSummary, TrackedChangesReview } from './TrackedChangesReview';
import { ambiguousReviewIds, decideReviewChange, resolveReviewAuthor, reviewRevisionIds, TrackedChangesContext, type TrackedChangesOptions } from './trackedChanges';
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

test('cards are one header row with icon-only by-id actions and active navigation', () => {
  const accept = mock((_id: number) => {}), reject = mock((_id: number) => {}), next = mock(() => {}), previous = mock(() => {});
  const range = mock(() => {});
  const view = render(<TrackedChangeCard {...sidebarProps} change={change} replies={[]} onAcceptById={accept} onRejectById={reject} onAccept={range}
    onNext={next} onPrevious={previous} index={1} total={3} />);
  const root = view.container.querySelector('.docx-tracked-change-card.docx-review-card')!;
  expect(root.getAttribute('data-active')).toBe('false');
  expect(root.getAttribute('data-kind')).toBe('human');
  expect(root.getAttribute('data-revision-id')).toBe('7');
  for (const part of ['name', 'meta', 'actions']) expect(root.querySelector(`.docx-review-card__header > .docx-review-card__${part}`)).not.toBeNull();
  expect(root.querySelector('.docx-review-card__diff, del, ins')).toBeNull();
  expect(root.textContent).not.toContain('weeks');
  const acceptButton = view.getByRole('button', { name: 'Accept change' }), rejectButton = view.getByRole('button', { name: 'Reject change' });
  expect([acceptButton.title, rejectButton.title]).toEqual(['Accept change', 'Reject change']);
  expect([acceptButton.textContent, rejectButton.textContent]).toEqual(['', '']);
  expect([...acceptButton.querySelectorAll('svg path')].map(path => path.getAttribute('d'))).toEqual(['M20 6 9 17l-5-5']);
  expect([...rejectButton.querySelectorAll('svg path')].map(path => path.getAttribute('d'))).toEqual(['M18 6 6 18', 'm6 6 12 12']);
  fireEvent.click(acceptButton);
  fireEvent.click(rejectButton);
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

test('insertions, deletions and structural revisions render the same minimal card', () => {
  const view = render(<TrackedChangeCard {...sidebarProps} change={{ ...change, type: 'insertion' }} replies={[]} />);
  for (const type of ['insertion', 'deletion', 'rowInserted'] as const) {
    view.rerender(<TrackedChangeCard {...sidebarProps} change={{ ...change, type }} replies={[]} />);
    expect(view.container.querySelector('.docx-review-card__diff, del, ins')).toBeNull();
    expect(view.getAllByRole('button').map(button => button.getAttribute('aria-label'))).toEqual(['Accept change', 'Reject change']);
  }
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

type HarnessRect = { pageIndex: number; x: number; y: number; width: number; height: number };
const onePage = [{ page: { pageIndex: 0 }, bounds: { x: 0, y: 0, width: 600, height: 800 } }];
const twoChanges = [change, { ...change, revisionId: 20, author: 'Riley', from: 30, to: 40 }];
function DotsHarness({ options = { display: 'dots' }, readOnly = false, pages = onePage, changeRects = [{ pageIndex: 0, x: 30, y: 100, width: 50, height: 12 }], changes = twoChanges }:
  { options?: TrackedChangesOptions; readOnly?: boolean; pages?: typeof onePage; changeRects?: HarnessRect[]; changes?: TrackedChangeEntry[] }) {
  const host = useRef<HTMLDivElement>(null), scroller = useRef<HTMLDivElement>(null);
  const [active, setActive] = useState<string | null>(null);
  const items = useCommentSidebarItems({ comments: [], trackedChanges: changes, callbacks: {} });
  const rects = new Map(items.map(item => [item.id, changeRects]));
  return <div ref={scroller}><div ref={host}><TrackedChangesReview items={items} options={options} pages={pages}
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
  await waitFor(() => expect(view.container.querySelectorAll('.docx-review-balloon')).toHaveLength(0));
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

test('a balloon below a change that continues onto the next page stays beside its line on the dot page', async () => {
  const pages = [onePage[0]!, { page: { pageIndex: 1 }, bounds: { x: 0, y: 820, width: 600, height: 800 } }];
  const view = render(<DotsHarness pages={pages} changeRects={[{ pageIndex: 0, x: 30, y: 760, width: 500, height: 12 }, { pageIndex: 0, x: 0, y: 774, width: 400, height: 12 },
    { pageIndex: 1, x: 0, y: 60, width: 200, height: 12 }]} />);
  const scroller = view.container.firstElementChild as HTMLDivElement;
  const host = scroller.firstElementChild as HTMLDivElement;
  scroller.getBoundingClientRect = () => ({ left: 0, top: 0, right: 700, bottom: 800, width: 700, height: 800 }) as DOMRect;
  host.getBoundingClientRect = () => ({ left: 0, top: 0, right: 700, bottom: 1700, width: 700, height: 1700 }) as DOMRect;
  Object.defineProperty(host, 'clientWidth', { value: 700 });
  fireEvent.resize(window);
  fireEvent.focus(await view.findByRole('button', { name: 'Change by Alex Reed' }));
  const top = Number.parseFloat(view.container.querySelector<HTMLElement>('.docx-review-balloon')!.style.top);
  expect(top).toBe(48 + 786 + 6);
  const leader = view.container.querySelector('svg.docx-review-leader path')!.getAttribute('d')!;
  expect(leader.startsWith(`M${50 + 200} ${48 + 820 + 72} `)).toBe(true);
});

test('dots drop a decided change on the next render without waiting for the layout pass', async () => {
  const view = render(<DotsHarness />);
  await waitFor(() => expect(view.container.querySelectorAll('.docx-review-dot')).toHaveLength(2));
  view.rerender(<DotsHarness changes={[twoChanges[1]!]} />);
  await waitFor(() => expect(view.container.querySelectorAll('.docx-review-dot')).toHaveLength(1));
  expect(view.container.querySelector('.docx-review-dot')?.getAttribute('aria-label')).toBe('Change by Riley');
});

test('a change row appears only where position alone cannot identify the change', () => {
  const line = (id: string, top: number, pageIndex = 0, ...more: number[]) =>
    ({ id, pageIndex, top, height: 16, lines: [top, ...more].map(lineTop => ({ pageIndex, top: lineTop, height: 16 })) });
  expect(ambiguousReviewIds([line('a', 100), line('b', 101), line('c', 200)], new Map([['a', 100], ['b', 140], ['c', 200]]))).toEqual(['a', 'b']);
  expect(ambiguousReviewIds([line('a', 100), line('b', 101, 1)])).toEqual([]);
  expect(ambiguousReviewIds([line('a', 100), line('b', 140)], new Map([['a', 96], ['b', 150]]))).toEqual([]);
  expect(ambiguousReviewIds([line('a', 100), line('b', 120)], new Map([['a', 100], ['b', 150]]))).toEqual(['b']);
  expect(ambiguousReviewIds([line('a', 100)], new Map([['a', 84]]))).toEqual([]);
  expect(ambiguousReviewIds([line('a', 100)], new Map([['a', 60]]))).toEqual(['a']);
  expect(ambiguousReviewIds([line('a', 100, 0, 120), line('b', 121)], new Map([['a', 100], ['b', 120]]))).toEqual(['a', 'b']);
  expect(ambiguousReviewIds([line('a', 100, 0, 120), line('b', 140)], new Map([['a', 100], ['b', 140]]))).toEqual([]);
});

test('ambiguous cards show the change in the document mark styling, old then new', () => {
  const card = (entry: TrackedChangeEntry, ambiguous: string[]) => <TrackedChangesContext.Provider value={{ ambiguous: new Set(ambiguous) }}>
    <TrackedChangeCard {...sidebarProps} change={entry} reviewItemId="r" replies={[]} /></TrackedChangesContext.Provider>;
  const view = render(card(change, ['r']));
  const row = view.container.querySelector('.docx-review-card__header + .docx-review-card__diff')!;
  expect(view.container.querySelector('.docx-review-card')?.getAttribute('data-ambiguous')).toBe('true');
  expect([...row.children].map(part => [part.tagName, part.className, part.textContent])).toEqual([
    ['DEL', 'docx-review-diff-del', 'six weeks'], ['INS', 'docx-review-diff-ins', 'four weeks']]);
  view.rerender(card({ ...change, type: 'insertion' }, ['r']));
  expect(view.container.querySelector('.docx-review-card__diff')?.innerHTML).toBe('<ins class="docx-review-diff-ins">four weeks</ins>');
  view.rerender(card({ ...change, type: 'deletion', text: 'six weeks' }, ['r']));
  expect(view.container.querySelector('.docx-review-card__diff')?.innerHTML).toBe('<del class="docx-review-diff-del">six weeks</del>');
  view.rerender(card(change, ['other']));
  expect(view.container.querySelector('.docx-review-card__diff, del, ins')).toBeNull();
  expect(view.container.querySelector('.docx-review-card')?.getAttribute('data-ambiguous')).toBe('false');
});

test('cards on one line both show their change, and hovering a card adds its leader beside the active one', async () => {
  const view = render(<DotsHarness options={{ display: 'cards' }} />);
  const cards = await waitFor(() => { const cards = view.container.querySelectorAll<HTMLElement>('.docx-review-card'); expect(cards).toHaveLength(2); return cards; });
  await waitFor(() => expect([...cards].map(card => card.querySelector('.docx-review-card__diff')?.textContent)).toEqual(['six weeksfour weeks', 'six weeksfour weeks']));
  const leaders = () => [...view.container.querySelectorAll('svg.docx-review-leader')].map(leader => leader.getAttribute('data-review-item'));
  expect(leaders()).toEqual([]);
  fireEvent.mouseOver(cards[1]!.querySelector('.docx-review-card__name')!);
  await waitFor(() => expect(leaders()).toHaveLength(1));
  const riley = leaders()[0];
  fireEvent.click(cards[0]!);
  await waitFor(() => expect(leaders()).toHaveLength(2));
  fireEvent.mouseLeave(view.container.firstElementChild!.firstElementChild!);
  await waitFor(() => expect(leaders()).toHaveLength(1));
  expect(leaders()[0]).not.toBe(riley);
});
