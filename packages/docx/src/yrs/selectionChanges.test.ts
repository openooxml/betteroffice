import { expect, spyOn, test } from 'bun:test';
import { createYrsSession, withYrsSelectionCause, yrsSelectionCause } from './index';

test('selection attribution follows the session version and a local setter supersedes remote repair', async () => {
  const session = await createYrsSession();
  const peer = await createYrsSession();
  try {
    const [paraId] = session.loadStories([{ storyId: 'body', paragraphs: [{ text: 'abcdef' }] }]).body!;
    const point = (offset: number) => ({ story: 'body', paraId: paraId!, offset });
    session.setSelection(point(1), point(3));
    session.setSelection(point(2));
    expect(yrsSelectionCause(session)).toBeUndefined();
    peer.applyUpdate(session.encodeState());
    peer.insertText(point(0), 'remote');
    session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
    expect(yrsSelectionCause(session)).toBe('remote');
    const moved = session.selection()!;
    withYrsSelectionCause(session, yrsSelectionCause(session), () => session.setSelection(moved.anchor, moved.head));
    expect(yrsSelectionCause(session)).toBe('remote');
    session.setSelection(point(2));
    expect(yrsSelectionCause(session)).toBeUndefined();
    withYrsSelectionCause(session, 'seed', () => session.setSelection(point(2)));
    expect(yrsSelectionCause(session)).toBe('seed');
    const selection = spyOn(session, 'selection').mockReturnValue(null);
    try {
      expect(yrsSelectionCause(session)).toBeUndefined();
      withYrsSelectionCause(session, 'seed', () => session.setSelection(point(2)));
      expect(yrsSelectionCause(session)).toBeUndefined();
    } finally { selection.mockRestore(); }
    expect(yrsSelectionCause(session)).toBe('seed');
    session.insertText(point(6), '!');
    expect(yrsSelectionCause(session)).toBeUndefined();
    session.setSelection(point(1));
    expect(yrsSelectionCause(session)).toBeUndefined();
  } finally { peer.destroy(); session.destroy(); }
});
