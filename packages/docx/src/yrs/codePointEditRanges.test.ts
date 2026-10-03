import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { findWordBoundaries } from '../utils/textSelection';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsLoc, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
let nextClientId = 54000;

function texts(session: YrsSession): string[] {
  return session.paragraphs('body').map((paragraph) => paragraph.text);
}

async function assertRoundTrips(session: YrsSession): Promise<void> {
  const peer = await createYrsSession({ clientId: nextClientId++ });
  try {
    peer.loadState(session.encodeState());
    expect(texts(peer)).toEqual(texts(session));
  } finally {
    peer.destroy();
  }
}

async function withStory(
  edit: (session: YrsSession, at: (offset: number) => YrsLoc) => void
): Promise<void> {
  const session = await createYrsSession({ clientId: nextClientId++ });
  try {
    const { paraId } = session.createStory('body', 'a😀b');
    const at = (offset: number): YrsLoc => ({ story: 'body', paraId, offset });
    session.splitParagraph(at(3));
    expect(texts(session)).toEqual(['a😀', 'b']);
    edit(session, at);
    await assertRoundTrips(session);
  } finally {
    session.destroy();
  }
}

describe('code point edit ranges', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('deletes the whole emoji selected by double-click and keeps both paragraphs', async () => {
    await withStory((session, at) => {
      const paragraph = session.paragraphs('body')[0];
      const [start, end] = findWordBoundaries(paragraph.text, 1);
      const receipt = session.deleteRange({ story: 'body', start: at(start), end: at(end) });
      expect(texts(session)).toEqual(['a', 'b']);
      const { paraId } = paragraph;
      expect(receipt.range).toEqual({
        story: 'body',
        start: { paraId, offset: 1 },
        end: { paraId, offset: 1 },
      });
    });
  });

  it('replaces the whole emoji selected by double-click and keeps both paragraphs', async () => {
    await withStory((session, at) => {
      const paragraph = session.paragraphs('body')[0];
      const [start, end] = findWordBoundaries(paragraph.text, 1);
      const receipt = session.replaceRange({ story: 'body', start: at(start), end: at(end) }, 'x');
      expect(texts(session)).toEqual(['ax', 'b']);
      expect(receipt.range).toEqual({
        story: 'body',
        start: { paraId: paragraph.paraId, offset: 1 },
        end: { paraId: paragraph.paraId, offset: 2 },
      });
    });
  });

  it('inserts text before the emoji when the caret is inside its surrogate pair', async () => {
    await withStory((session, at) => {
      const { paraId } = at(2);
      const receipt = session.insertText(at(2), 'x');
      expect(texts(session)).toEqual(['ax😀', 'b']);
      expect(receipt.range).toEqual({
        story: 'body',
        start: { paraId, offset: 1 },
        end: { paraId, offset: 2 },
      });
    });
  });

  it('reports an image inserted inside a surrogate pair before the emoji', async () => {
    await withStory((session, at) => {
      const { paraId } = at(2);
      const receipt = session.insertImage(at(2), { rId: 'rIdImage' });
      expect(texts(session)).toEqual(['a😀', 'b']);
      expect(receipt.range).toEqual({
        story: 'body',
        start: { paraId, offset: 1 },
        end: { paraId, offset: 2 },
      });
    });
  });
});
