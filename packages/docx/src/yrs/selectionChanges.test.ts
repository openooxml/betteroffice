import { expect, spyOn, test } from 'bun:test';
import { createYrsSession, subscribeYrsSelection, withYrsSelectionCause, yrsSelectionCause } from './index';

test('session selection subscribers receive caret moves and remote rebases without formatting notifications', async () => {
  const session = await createYrsSession();
  const peer = await createYrsSession();
  let stop = () => {};
  try {
    const [paraId] = session.loadStories([{ storyId: 'body', paragraphs: [{ text: 'abcdef' }] }]).body!;
    const point = (offset: number) => ({ story: 'body', paraId: paraId!, offset });
    session.setSelection(point(1));
    const changes: Array<{ offset: number; cause: ReturnType<typeof yrsSelectionCause> }> = [];
    stop = subscribeYrsSelection(session, () => changes.push({ offset: session.selection()!.head.offset,
      cause: yrsSelectionCause(session) }));
    session.setSelection(point(2));
    expect(changes).toEqual([{ offset: 2, cause: undefined }]);
    session.setSelection(point(2));
    session.formatRange({ story: 'body', start: point(0), end: point(6) }, { bold: true });
    expect(changes).toEqual([{ offset: 2, cause: undefined }]);
    peer.loadState(session.encodeState());
    peer.insertText(point(0), 'remote');
    session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
    expect(changes).toEqual([{ offset: 2, cause: undefined }, { offset: 8, cause: 'remote' }]);
    stop();
    session.setSelection(point(3));
    expect(changes).toHaveLength(2);
  } finally { stop(); peer.destroy(); session.destroy(); }
});

test('removing the selected story publishes an empty remote selection', async () => {
  const session = await createYrsSession();
  const peer = await createYrsSession();
  let stop = () => {};
  try {
    const [paraId] = session.loadStories([{ storyId: 'hf:rIdH1', paragraphs: [{ text: 'Header' }] }])['hf:rIdH1']!;
    session.setSelection({ story: 'hf:rIdH1', paraId: paraId!, offset: 2 });
    peer.loadState(session.encodeState());
    expect(peer.hasStory('hf:rIdH1')).toBe(true);
    const changes: Array<{ selection: ReturnType<typeof session.selection>; cause: ReturnType<typeof yrsSelectionCause> }> = [];
    stop = subscribeYrsSelection(session, () => changes.push({ selection: session.selection(), cause: yrsSelectionCause(session) }));
    peer.deleteStory('hf:rIdH1');
    session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
    expect(session.selection()).toBeNull();
    expect(changes).toEqual([{ selection: null, cause: 'remote' }]);
  } finally { stop(); peer.destroy(); session.destroy(); }
});

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
