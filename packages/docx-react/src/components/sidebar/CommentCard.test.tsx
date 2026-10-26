import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import type { Comment } from '@betteroffice/docx/types/content';
import type { CollaborationPeer } from '@betteroffice/docx/collaboration';
import { CommentCard } from './CommentCard';
import { AddCommentCard } from './AddCommentCard';
import { ResolvedCommentMarker } from './ResolvedCommentMarker';
import { TrackedChangeCard } from './TrackedChangeCard';
import { TrackedChangesContext } from './trackedChanges';
import type { TrackedChangeEntry } from './cardUtils';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>(resolve => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const text = (value: string) => [{ type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: value }] }] }] as unknown as Comment['content'];
const comment: Comment = { id: 1, author: 'Mia Chen', date: '2026-01-01T09:00:00Z', content: text('Name the owner') };
const replies: Comment[] = [
  { id: 2, author: 'Ravi Patel', date: '2026-01-02T09:00:00Z', parentId: 1, content: text('First reply') },
  { id: 3, author: 'Lena Fischer', date: '2026-01-03T09:00:00Z', parentId: 1, content: text('Second reply') },
];
const change: TrackedChangeEntry = { revisionId: 7, type: 'insertion', author: 'Mia Chen', date: '2026-10-01T09:00:00Z', text: 'four', from: 1, to: 5 };
const sidebarProps = { onToggleExpand() {}, measureRef() {} };
const css = readFileSync(new URL('../../../../docx/src/styles/editor.css', import.meta.url), 'utf8');
const rule = (selector: string) => css.split('}').find(block => block.split('{')[0]!.split(',').some(part => part.trim() === selector)) ?? '';

test('comment cards use the shared review card chrome and header layout of tracked change cards', () => {
  const view = render(<>
    <CommentCard {...sidebarProps} isExpanded comment={comment} replies={replies} />
    <TrackedChangeCard {...sidebarProps} isExpanded={false} change={change} replies={[]} />
  </>);
  const card = view.container.querySelector<HTMLElement>('.docx-comment-card')!;
  const tracked = view.container.querySelector<HTMLElement>('.docx-tracked-change-card')!;
  expect(card.classList.contains('docx-review-card')).toBe(true);
  expect(card.getAttribute('data-active')).toBe('true');
  for (const part of ['name', 'meta', 'actions']) expect(card.querySelector(`.docx-review-card__header > .docx-review-card__${part}`)).not.toBeNull();
  expect(card.querySelector('.docx-review-card__header > .docx-review-avatar')?.textContent).toBe('MC');
  expect(card.style.getPropertyValue('--doc-review-author')).toBe(tracked.style.getPropertyValue('--doc-review-author'));
  expect(card.getAttribute('style')?.replace(/\s/g, '')).toMatch(/^--doc-review-author:[^;]+;?$/);
  expect(card.querySelector('.docx-review-card__body')?.textContent).toBe('Name the owner');
  expect(view.getByRole('button', { name: 'Resolve' }).className).toBe('docx-review-resolve');
  expect(view.getByRole('button', { name: 'More options' }).className).toBe('docx-review-more');
  expect(card.querySelectorAll('.docx-review-reply .docx-review-avatar')).toHaveLength(2);
  expect(card.querySelector('.docx-review-card__replies .docx-review-composer .docx-review-input')).not.toBeNull();
  expect(card.querySelectorAll('[style]').length).toBe(card.querySelectorAll('.docx-review-avatar').length);
});

test('the comment classes are styled from the review tokens, sharing the tracked change button rules', () => {
  expect(rule('.docx-review-button')).toBe(rule('.docx-review-accept-all'));
  expect(rule('.docx-review-resolve')).toBe(rule('.docx-review-accept'));
  expect(rule('.docx-review-more')).toBe(rule('.docx-review-reject'));
  for (const selector of ['.docx-review-card__body', '.docx-review-reply', '.docx-review-input', '.docx-review-menu', '.docx-review-menu__item', '.docx-review-resolved-marker']) {
    const body = rule(selector);
    expect(body).toContain('var(--doc-review-');
    expect(body).not.toMatch(/#[0-9a-f]{3,8}\b|rgba?\(/i);
  }
});

test('collapsed cards hide actions and show only the latest reply with a count', () => {
  const view = render(<CommentCard {...sidebarProps} isExpanded={false} comment={comment} replies={replies} />);
  expect(view.container.querySelector('.docx-review-card__actions')).toBeNull();
  expect(view.container.querySelector('.docx-review-input')).toBeNull();
  expect(view.container.querySelector('.docx-review-thread__more')?.textContent).toBe('1 reply');
  const shown = view.container.querySelectorAll('.docx-review-reply__body');
  expect([...shown].map(node => [node.textContent, node.getAttribute('data-clamped')])).toEqual([['Second reply', 'true']]);
});

test('resolve, menu delete, escape and reply keep working', () => {
  const onResolve = mock((_id: number) => {}), onDelete = mock((_id: number) => {}), onReply = mock((_id: number, _text: string) => {});
  const toggle = mock(() => {});
  const view = render(<CommentCard {...sidebarProps} onToggleExpand={toggle} isExpanded comment={comment} replies={[]}
    onResolve={onResolve} onDelete={onDelete} onReply={onReply} />);
  fireEvent.click(view.getByRole('button', { name: 'Resolve' }));
  expect(onResolve).toHaveBeenCalledWith(1);
  const more = view.getByRole('button', { name: 'More options' });
  fireEvent.click(more);
  expect(more.getAttribute('aria-expanded')).toBe('true');
  expect(view.getByRole('menu').className).toBe('docx-review-menu');
  fireEvent.keyDown(document, { key: 'Escape' });
  expect(view.queryByRole('menu')).toBeNull();
  fireEvent.click(more);
  fireEvent.click(view.getByRole('menuitem', { name: 'Delete' }));
  expect(onDelete).toHaveBeenCalledWith(1);
  expect(view.queryByRole('menu')).toBeNull();
  fireEvent.click(view.getByPlaceholderText('Reply or add others with @'));
  const input = view.getByPlaceholderText('Reply or add others with @');
  expect(view.getByRole('button', { name: 'Reply' }).hasAttribute('disabled')).toBe(true);
  fireEvent.change(input, { target: { value: '  A reply ' } });
  expect(view.getByRole('button', { name: 'Reply' }).className).toBe('docx-review-button docx-review-button--primary');
  fireEvent.keyDown(input, { key: 'Enter' });
  expect(onReply).toHaveBeenCalledWith(1, 'A reply');
  expect(toggle).not.toHaveBeenCalled();
});

test('resolved cards show the status, reopen, and no reply input', () => {
  const onUnresolve = mock((_id: number) => {});
  const view = render(<CommentCard {...sidebarProps} isExpanded comment={{ ...comment, done: true }} replies={[]} onUnresolve={onUnresolve} />);
  const card = view.container.querySelector('.docx-comment-card')!;
  expect(card.getAttribute('data-resolved')).toBe('true');
  expect(card.querySelector('.docx-review-card__meta .docx-review-card__status')?.textContent).toBe('Resolved');
  expect(card.querySelector('.docx-review-input')).toBeNull();
  fireEvent.click(view.getByRole('button', { name: 'Reopen' }));
  expect(onUnresolve).toHaveBeenCalledWith(1);
  const marker = render(<ResolvedCommentMarker {...sidebarProps} isExpanded={false} comment={{ ...comment, done: true }} />);
  expect(marker.container.querySelector('[data-comment-resolved="true"]')?.className).toBe('docx-review-resolved-marker');
});

test('comment authors take collaboration colours and host appearance like tracked changes', () => {
  const peer = { clientId: 3, user: { name: 'Mia Chen', color: '#123456' } } as CollaborationPeer;
  const view = render(<TrackedChangesContext.Provider value={{ peers: [peer], options: { getAuthorAppearance: ({ name }) => name === 'Ravi Patel' ? { color: '#abcdef' } : undefined } }}>
    <CommentCard {...sidebarProps} isExpanded comment={comment} replies={replies.slice(0, 1)} />
  </TrackedChangesContext.Provider>);
  const avatars = [...view.container.querySelectorAll<HTMLElement>('.docx-review-avatar')];
  expect(avatars.map(avatar => avatar.style.getPropertyValue('--doc-review-author'))).toEqual(['#123456', '#abcdef']);
});

test('the add comment composer uses the review card, input and buttons', async () => {
  const onSubmit = mock((_text: string) => {});
  const view = render(<AddCommentCard {...sidebarProps} isExpanded onSubmit={onSubmit} />);
  expect(view.container.querySelector('.docx-comment-composer.docx-review-card textarea.docx-review-input')).not.toBeNull();
  expect(view.getByRole('button', { name: 'Cancel' }).className).toBe('docx-review-button');
  fireEvent.change(view.getByPlaceholderText('Add a comment...'), { target: { value: 'New note' } });
  await act(async () => { fireEvent.click(view.getByRole('button', { name: 'Comment' })); });
  expect(onSubmit).toHaveBeenCalledWith('New note');
});
