import { describe, expect, test } from 'bun:test';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { performYrsHistoryAction } from './yrsCommands';

describe('performYrsHistoryAction', () => {
  test('reports every story the undo changed instead of the active selection story', () => {
    const session = {
      historyStories: () => ['body', 'hf:rId7'],
      selection: () => ({
        anchor: { story: 'fn:2', paraId: 'note', offset: 0 },
        head: { story: 'fn:2', paraId: 'note', offset: 0 },
      }),
      undo: () => true,
    } as unknown as YrsSession;

    expect(performYrsHistoryAction(session, false)).toEqual({
      changed: true,
      stories: ['body', 'hf:rId7'],
    });
  });

  test('does not report a dirty story when history is unchanged', () => {
    const session = {
      selection: () => null,
      redo: () => false,
    } as unknown as YrsSession;

    expect(performYrsHistoryAction(session, true)).toEqual({ changed: false, stories: [] });
  });
});
