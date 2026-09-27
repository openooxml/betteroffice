import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  applyYrsToolbarFormatting,
  currentYrsToolbarSelection,
  storedYrsToolbarFormatting,
  withStoredYrsFormatting,
} from './yrsToolbar';

const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

async function selected(text: string, start: number, end: number) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', text);
  session.setSelection(
    { story: 'body', paraId, offset: start },
    { story: 'body', paraId, offset: end }
  );
  const map = () =>
    createYrsInputPositionMap(
      'body',
      session.paragraphs('body').map((paragraph) => ({
        paraId: paragraph.paraId,
        length: paragraph.text.length,
      }))
    );
  return { session, map, selection: () => currentYrsToolbarSelection(session, map())! };
}

describe('toolbar formatting adapter', () => {
  test('toggles superscript and subscript exclusively over a range', async () => {
    const { session, map, selection } = await selected('hello world', 0, 5);
    expect(applyYrsToolbarFormatting(session, map(), 'superscript')).toBe(true);
    expect(selection().context.superscript).toBe(true);
    expect(selection().context.subscript).toBe(false);
    expect(applyYrsToolbarFormatting(session, map(), 'subscript')).toBe(true);
    expect(selection().context.superscript).toBe(false);
    expect(selection().context.subscript).toBe(true);
    session.setSelection(
      { story: 'body', paraId: selection().context.paraId, offset: 0 },
      { story: 'body', paraId: selection().context.paraId, offset: 11 }
    );
    expect(selection().context.subscript).toBe('mixed');
  });

  test('reads highlight from the authoritative selection', async () => {
    const { session, map, selection } = await selected('hello world', 0, 5);
    applyYrsToolbarFormatting(session, map(), { type: 'highlightColor', value: 'FFFF00' });
    expect(selection().context.highlight).toBe('yellow');
    applyYrsToolbarFormatting(session, map(), { type: 'highlightColor', value: 'none' });
    expect(selection().context.highlight).toBeNull();
  });

  test('stores script and highlight for a collapsed caret and overlays them', async () => {
    const { selection } = await selected('hello', 5, 5);
    const context = selection().context;
    expect(storedYrsToolbarFormatting(context, 'superscript')).toEqual({
      type: 'toggle',
      mark: 'superscript',
      active: false,
    });
    const overlaid = withStoredYrsFormatting(selection(), {
      clear: false,
      delta: { highlight: 'green', other: { superscript: true, subscript: null } },
    });
    expect(overlaid.context.superscript).toBe(true);
    expect(overlaid.context.subscript).toBe(false);
    expect(overlaid.context.highlight).toBe('green');
    const cleared = withStoredYrsFormatting(selection(), { clear: true, delta: {} });
    expect(cleared.context.superscript).toBe(false);
    expect(cleared.context.highlight).toBeNull();
  });
});
