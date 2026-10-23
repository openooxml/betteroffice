import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import type { DocxReviewTarget } from '@betteroffice/docx/yrs';
import { DocxAsyncOnlyError } from '../hooks/useDocxEditorRefApi';
import { resetDeprecatedViewerMembersForTests } from '../internals/deprecatedViewerMembers';
import { createPeerlessShellController, PeerlessSyncApiError } from './shellController';
import { proposalHost } from './workerProposalHost.testing';

afterEach(() => {
  mock.restore();
  resetDeprecatedViewerMembersForTests();
});

function viewer(readOnly = true) {
  spyOn(console, 'warn').mockImplementation(() => {});
  return createPeerlessShellController({ read: async () => ({} as never), focus() {}, zoom: 1,
    readOnly, mode: readOnly ? 'editing' : 'viewing' });
}

test('worker viewer synchronous reads and comment mutations return null', () => {
  for (const readOnly of [true, false]) {
    const { ref } = viewer(readOnly);
    expect(ref.getEditorRef()).toBeNull();
    expect(ref.getSelectionInfo()).toBeNull();
    expect(ref.getPositionAtPoint(10, 20)).toBeNull();
    expect(ref.addComment({ paraId: 'p', text: 'Comment', author: 'Author' })).toBeNull();
    expect(ref.replyToComment(1, 'Reply', 'Author')).toBeNull();
  }
});

test('worker viewer synchronous formatting mutations return false', () => {
  for (const readOnly of [true, false]) {
    const { ref } = viewer(readOnly);
    expect(ref.applyFormatting({ paraId: 'p', marks: { bold: true } })).toBe(false);
    expect(ref.setParagraphStyle({ paraId: 'p', styleId: 'Normal' })).toBe(false);
    expect(ref.insertBreak({ paraId: 'p', type: 'page' })).toBe(false);
  }
});

for (const options of [{ readOnly: true }, { mode: 'viewing' as const }]) {
  test(`worker viewer async comment mutations return null and resolveComment is a no-op with ${options.readOnly ? 'readOnly' : 'viewing'}`, async () => {
    const h = proposalHost(options);
    const read = mock(async () => ({} as never));
    const onCommentsChange = mock(() => {});
    h.controller.update({ ...h.settings, read, onCommentsChange });
    try {
      const comment = h.controller.ref.insertComment({ paraId: 'p', text: 'Comment', author: 'Author' });
      const reply = h.controller.ref.insertCommentReply(1, 'Reply', 'Author');
      expect(comment).toBeInstanceOf(Promise);
      expect(reply).toBeInstanceOf(Promise);
      expect(await comment).toBeNull();
      expect(await reply).toBeNull();
      expect(h.controller.ref.resolveComment(1)).toBeUndefined();
      expect(h.submit).not.toHaveBeenCalled();
      expect(read).not.toHaveBeenCalled();
      expect(h.readAt).not.toHaveBeenCalled();
      expect(onCommentsChange).not.toHaveBeenCalled();
    } finally {
      h.controller.dispose();
    }
  });
}

test('worker viewer synchronous document reads throw DocxAsyncOnlyError with twins', () => {
  for (const readOnly of [true, false]) {
    const { ref } = viewer(readOnly);
    for (const [member, args, twin] of [
      ['getDocument', [], 'readParagraphs or exportStructuredWithPages'],
      ['getPageContent', [1], 'exportStructuredWithPages'],
      ['findInDocument', ['text'], 'findParagraphs'],
    ] as const) {
      try {
        Reflect.apply(ref[member], ref, args);
        throw new Error('Expected an async-only error');
      } catch (error) {
        expect(error).toBeInstanceOf(DocxAsyncOnlyError);
        expect(error).toMatchObject({ member, use: twin });
      }
    }
  }
});

test('worker viewer synchronous navigation starts async twins and returns true', async () => {
  const controller = viewer();
  const paragraph = mock(async (_paraId: string) => { throw new Error('gone'); });
  const review = mock(async (_target: DocxReviewTarget) => false);
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, readOnly: true,
    scrollToParagraph: paragraph, scrollToReviewTarget: review });
  expect(controller.ref.scrollToParaId('p')).toBe(true);
  expect(controller.ref.scrollToCommentId(1)).toBe(true);
  expect(controller.ref.scrollToChangeId(2)).toBe(true);
  expect(paragraph).toHaveBeenCalledWith('p');
  expect(review.mock.calls).toEqual([[{ kind: 'comment', id: 1 }], [{ kind: 'revision', revisionId: '2' }]]);
  await Promise.resolve();
});

test('worker viewer content listeners never fire and unsubscribe is a no-op', () => {
  const { ref } = viewer();
  const listener = mock(() => {});
  const unsubscribe = ref.onContentChange(listener);
  expect(unsubscribe()).toBeUndefined();
  expect(unsubscribe()).toBeUndefined();
  expect(listener).not.toHaveBeenCalled();
});

test('worker viewer refusals follow mode updates without enabling synchronous editing', () => {
  const controller = viewer();
  expect(controller.ref.getSelectionInfo()).toBeNull();
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1 });
  expect(() => controller.ref.getSelectionInfo()).toThrow(PeerlessSyncApiError);
});

for (const [member, twin, args] of [
  ['scrollToPosition', 'scrollToParagraph', [10]],
  ['highlightRange', 'readPositionAtPoint', [10, 24]],
] as const) {
  test(`worker viewer ${member} warns once with ${twin} and retains its sync refusal`, () => {
    const { ref } = viewer();
    for (let call = 0; call < 2; call++) expect(() => Reflect.apply(ref[member] as Function, ref, args)).toThrow(PeerlessSyncApiError);
    expect(console.warn).toHaveBeenCalledTimes(1);
    expect(console.warn).toHaveBeenCalledWith(expect.stringContaining(`Use ${twin}.`));
    expect(console.warn).toHaveBeenCalledWith(expect.stringContaining('throws DocxSyncApiError in worker-only viewer sessions'));
  });
}

test('worker viewer point reads warn once that they return null', () => {
  const { ref } = viewer();
  expect(ref.getPositionAtPoint(10, 20)).toBeNull();
  expect(ref.getPositionAtPoint(10, 20)).toBeNull();
  expect(console.warn).toHaveBeenCalledTimes(1);
  expect(console.warn).toHaveBeenCalledWith(expect.stringContaining('returns null in worker-only viewer sessions. Use readPositionAtPoint.'));
});

test('worker viewer synchronous proposals return false without host permission', () => {
  const { ref } = viewer();
  expect(ref.proposeChange({ paraId: 'p', search: 'old', replaceWith: 'new', author: 'Author' })).toBe(false);
});

test('worker viewer synchronous proposals queue allowed changes behind pending input', async () => {
  const h = proposalHost({ readOnly: true });
  let submitted!: () => void;
  const admission = new Promise<void>(resolve => { submitted = resolve; });
  const submit = h.binding.commands.submit.bind(h.binding.commands);
  h.binding.commands.submit = (...args) => { const result = submit(...args); submitted(); return result; };
  const read = mock(async () => ({ status: 'ok', version: 'native-0', value: { sessionId: 's', paragraphs: [
    { session: { story: 'body', paraId: 'p' } },
  ] } }));
  const readAt = h.settings.readAt;
  h.controller.update({ ...h.settings, read: read as never, readAt: (async (request: { kind: string }, version: string) =>
    request.kind === 'paragraphIdentities' ? { status: 'ok', version, value: { sessionId: 's', paragraphs: [
      { session: { story: 'body', paraId: 'p' } },
    ] } } : readAt!(request as never, version)) as never });
  const composition = h.binding.admission.reserveComposition();
  expect(h.controller.ref.proposeChange({ paraId: 'p', search: 'old', replaceWith: 'new', author: 'Author' })).toBe(true);
  expect(read).not.toHaveBeenCalled();
  expect(h.submit).not.toHaveBeenCalled();
  h.binding.admission.settleComposition(composition.slot, null);
  await admission;
  expect(await h.controller.ref.getProposals()).toMatchObject({ proposals: [{ state: 'proposed' }] });
  expect(h.submit.mock.calls[0]).toMatchObject(['propose', { expectVersion: 'native-0', proposals: [{
    paragraph: { kind: 'session', sessionId: 's', story: 'body', paraId: 'p' },
    op: 'replaceText', search: 'old', replaceWith: 'new', suggest: { author: 'Author' },
  }] }, { fence: true }]);
  h.controller.dispose();
});

for (const [reason, message] of [
  ['missing-target', 'missing-target'],
  ['proposal-id-conflict', 'proposal-id-conflict'],
  ['stale-version', 'the document changed since the expected version was read'],
] as const) {
  test(`worker viewer queued proposals report ${reason} through onError`, async () => {
    const h = proposalHost({ readOnly: true });
    const onError = mock((_error: Error) => {});
    const readAt = h.settings.readAt;
    h.controller.update({ ...h.settings, onError, readAt: (async (request: { kind: string }, version: string) =>
      request.kind === 'paragraphIdentities' ? { status: 'ok', version, value: { sessionId: 's', paragraphs:
        reason === 'missing-target' ? [] : [{ session: { story: 'body', paraId: 'p' } }],
      } } : readAt!(request as never, version)) as never });
    if (reason !== 'missing-target') h.refuse(reason);
    try {
      expect(h.controller.ref.proposeChange({ paraId: 'p', search: 'old', replaceWith: 'new', author: 'Author' })).toBe(true);
      expect(await h.controller.ref.getProposals()).toMatchObject({ proposals: [] });
      expect(onError).toHaveBeenCalledTimes(1);
      expect(onError.mock.calls[0]![0]).toBeInstanceOf(Error);
      expect(onError.mock.calls[0]![0].message).toBe(message);
      expect(h.submit).toHaveBeenCalledTimes(reason === 'missing-target' ? 0 : 1);
    } finally {
      h.controller.dispose();
    }
  });
}
