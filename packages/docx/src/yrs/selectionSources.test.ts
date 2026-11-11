import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, test } from 'bun:test';

/**
 * Every place a selection enters editor command execution, and how it passes `laidOutStory` (the
 * root projection lays its story out) before a command uses it. A selection that fails is moved
 * next to the unavailable table that hides it (`besideHiddenStory`), never used as given.
 */
const SOURCES: Record<string, Record<string, { count: number; via: string }>> = {
  predecessorSelection: {
    'residentCommandAnchors.ts': { count: 2, via: 'resolveCommandSelection hands it to readCommandSelection' },
    'residentCommandSelection.ts': { count: 1, via: 'handed to readCommandSelection' },
    'residentCommandHost.ts': { count: 1, via: 'carries the retained selection into the command context' },
    'residentEngineWorker.ts': { count: 2, via: 'the prefix check only decides reuse; the vertical move focus is checked with the engine read storyLaidOut' },
  },
  'engine.selection()': {
    'residentCommandActions.ts': { count: 3, via: 'undo/redo: table ancestry before the undo, and the selection after it, checked with laidOutStory; hooks: the initial confirmed selection, read through readCommandSelection on every command' },
    'residentEngineWorker.ts': { count: 1, via: 'the vertical move focus, checked with the engine read storyLaidOut' },
  },
  '.cellSelection()': {
    'residentCommandTableSelection.ts': { count: 1, via: 'readCellSelection: the engine refuses a cell selection on an unavailable table, read as none' },
    'residentEngineWorker.ts': { count: 1, via: 'a document read of the cell selection, refused coded on an unavailable table' },
  },
};

const EXCLUDED = new Set(['residentCommandProtocol.ts', 'residentCommandTestSupport.ts']);

function occurrences(text: string): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const file of readdirSync(import.meta.dir)) {
    if (!file.endsWith('.ts') || file.endsWith('.test.ts') || EXCLUDED.has(file)) continue;
    const count = readFileSync(join(import.meta.dir, file), 'utf8').split(text).length - 1;
    if (count > 0) counts[file] = count;
  }
  return counts;
}

for (const [text, files] of Object.entries(SOURCES)) {
  test(`every ${text} source is listed with the way it passes the laid-out check`, () => {
    const expected = Object.fromEntries(Object.entries(files).map(([file, { count }]) => [file, count]));
    expect(occurrences(text)).toEqual(expected);
  });
}

test('the one door: every range resolution asks whether the layout shows the story, and a resolved command selection in a hidden one moves beside the table', () => {
  const anchors = readFileSync(join(import.meta.dir, 'residentCommandAnchors.ts'), 'utf8');
  const resolve = anchors.slice(anchors.indexOf('function selectionRange('), anchors.indexOf('export function readCommandSelection'));
  expect(resolve).toContain('if (!shown(engine, anchor.story))');
  const read = anchors.slice(anchors.indexOf('export function readCommandSelection'), anchors.indexOf('export function currentSelectionRange'));
  expect(read).toContain("if (range === 'hidden')");
  expect(read).toContain('besideHiddenStory(engine, anchor.story)');
});
