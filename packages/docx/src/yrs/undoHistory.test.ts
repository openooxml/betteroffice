import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsLoc, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

function endOf(session: YrsSession, story: string): YrsLoc {
  const paragraph = session.paragraphs(story)[0];
  return { story, paraId: paragraph.paraId, offset: paragraph.text.length };
}

function text(session: YrsSession, story: string): string {
  return session.paragraphs(story)[0].text;
}

describe('session undo history', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('keeps one ordered history across stories', async () => {
    const session = await createYrsSession({ clientId: 53005 });
    try {
      session.createStory('body', 'body');
      session.createStory('hf:rId7', 'header');
      session.setSelection(endOf(session, 'body'));
      session.insertText(endOf(session, 'body'), '!');
      session.setSelection(endOf(session, 'hf:rId7'));
      session.insertText(endOf(session, 'hf:rId7'), '?');

      expect(session.undo()).toBe(true);
      expect(session.historyStories()).toEqual(['hf:rId7']);
      expect(text(session, 'hf:rId7')).toBe('header');
      expect(text(session, 'body')).toBe('body!');

      expect(session.undo()).toBe(true);
      expect(session.historyStories()).toEqual(['body']);
      expect(text(session, 'body')).toBe('body');
      expect(session.undo()).toBe(false);
    } finally {
      session.destroy();
    }
  });

  it('starts a new step when a direct operation targets another story', async () => {
    const session = await createYrsSession({ clientId: 53007 });
    try {
      session.createStory('body', 'body');
      session.createStory('fn:2', 'note');
      session.insertText(endOf(session, 'body'), '!');
      session.insertText(endOf(session, 'fn:2'), '?');

      expect(session.undo()).toBe(true);
      expect(session.historyStories()).toEqual(['fn:2']);
      expect(text(session, 'body')).toBe('body!');
      expect(text(session, 'fn:2')).toBe('note');
    } finally {
      session.destroy();
    }
  });

  it('separates host actions while retaining sticky anchors through undo and redo', async () => {
    const session = await createYrsSession({ clientId: 53008 });
    try {
      const { paraId } = session.createStory('body', 'Antes … depois');
      const at = (offset: number): YrsLoc => ({ story: 'body', paraId, offset });
      const marker = session.encodeStickyPosition(at(6));
      session.beginUndoCapture();
      session.insertText(at(0), 'Prefixo ');
      session.stopUndoCapture();
      session.deleteRange({ story: 'body', start: at(14), end: at(15) });
      expect(text(session, 'body')).toBe('Prefixo Antes  depois');
      for (let cycle = 0; cycle < 3; cycle += 1) {
        expect(session.undo()).toBe(true);
        expect(text(session, 'body')).toBe('Prefixo Antes … depois');
        expect(session.resolveStickyPosition(marker)).toEqual(at(14));
        expect(session.historyStories()).toEqual(['body']);
        expect(session.redo()).toBe(true);
        expect(text(session, 'body')).toBe('Prefixo Antes  depois');
      }
      expect(session.undo()).toBe(true);
      expect(session.undo()).toBe(true);
      expect(text(session, 'body')).toBe('Antes … depois');
      expect(session.undo()).toBe(false);
    } finally {
      session.destroy();
    }
  });

  it('does not create empty steps or clear redo when capture is stopped repeatedly', async () => {
    const session = await createYrsSession({ clientId: 53009 });
    try {
      session.stopUndoCapture();
      session.createStory('body', 'body');
      session.beginUndoCapture();
      session.stopUndoCapture();
      session.stopUndoCapture();
      expect(session.canUndo()).toBe(false);
      session.insertText(endOf(session, 'body'), '!');
      session.stopUndoCapture();
      session.stopUndoCapture();
      expect(session.undo()).toBe(true);
      expect(text(session, 'body')).toBe('body');
      expect(session.undo()).toBe(false);
      session.stopUndoCapture();
      expect(session.canRedo()).toBe(true);
      expect(session.redo()).toBe(true);
      expect(text(session, 'body')).toBe('body!');
    } finally {
      session.destroy();
    }
  });

  it('groups keystrokes inside the capture window into one step', async () => {
    const session = await createYrsSession({ clientId: 53006 });
    try {
      session.createStory('body', 'body');
      session.insertText(endOf(session, 'body'), 'a');
      session.insertText(endOf(session, 'body'), 'b');
      await Bun.sleep(600);
      session.insertText(endOf(session, 'body'), 'c');

      expect(session.undo()).toBe(true);
      expect(text(session, 'body')).toBe('bodyab');
      expect(session.undo()).toBe(true);
      expect(text(session, 'body')).toBe('body');
    } finally {
      session.destroy();
    }
  });
});
