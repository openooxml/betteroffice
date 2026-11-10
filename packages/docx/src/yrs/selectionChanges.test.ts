import { expect, test } from 'bun:test';
import { createYrsSession, subscribeYrsSelection, withYrsSelectionCause, yrsSelectionCause } from './index';

test('raw selection setters notify and a local setter supersedes remote repair attribution', async () => {
  const session = await createYrsSession();
  const peer = await createYrsSession();
  const causes: Array<'remote' | 'seed' | undefined> = [];
  let detach = () => {};
  try {
    const [paraId] = session.loadStories([{ storyId: 'body', paragraphs: [{ text: 'abcdef' }] }]).body!;
    const point = (offset: number) => ({ story: 'body', paraId: paraId!, offset });
    session.setSelection(point(1), point(3));
    detach = subscribeYrsSelection(session, cause => causes.push(cause));
    session.setSelection(point(2));
    expect(causes).toEqual([undefined]);
    peer.applyUpdate(session.encodeState());
    peer.insertText(point(0), 'remote');
    session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
    expect(causes.at(-1)).toBe('remote');
    expect(yrsSelectionCause(session)).toBe('remote');
    const moved = session.selection()!;
    withYrsSelectionCause(session, yrsSelectionCause(session), () => session.setSelection(moved.anchor, moved.head));
    expect(causes.at(-1)).toBe('remote');
    session.setSelection(point(2));
    expect(causes.at(-1)).toBeUndefined();
    expect(yrsSelectionCause(session)).toBeUndefined();
    withYrsSelectionCause(session, 'seed', () => session.setSelection(point(2)));
    expect(causes.at(-1)).toBe('seed');
    session.insertText(point(6), '!');
    expect(yrsSelectionCause(session)).toBeUndefined();
    detach();
    const count = causes.length;
    session.setSelection(point(1));
    expect(causes).toHaveLength(count);
  } finally { detach(); peer.destroy(); session.destroy(); }
});
