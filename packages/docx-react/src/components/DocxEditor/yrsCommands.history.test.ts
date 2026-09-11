import { describe, expect, test } from 'bun:test';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { performYrsHistoryAction } from './yrsCommands';

describe('performYrsHistoryAction', () => {
  test('reports the mutated history story instead of the active selection story', () => {
    const session = {
      historyStory: () => 'body',
      selection: () => ({
        anchor: { story: 'fn:2', paraId: 'note', offset: 0 },
        head: { story: 'fn:2', paraId: 'note', offset: 0 },
      }),
      undo: () => true,
    } as unknown as YrsSession;

    expect(performYrsHistoryAction(session, false)).toEqual({ changed: true, story: 'body' });
  });

  test('does not report a dirty story when history is unchanged', () => {
    const session = {
      historyStory: () => 'body',
      selection: () => null,
      redo: () => false,
    } as unknown as YrsSession;

    expect(performYrsHistoryAction(session, true)).toEqual({ changed: false, story: null });
  });
});
