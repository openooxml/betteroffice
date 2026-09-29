import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

it('hasStory answers as storyIds().includes does', async () => {
  const session = await createYrsSession({ clientId: 76101 });
  try {
    session.createStory('body', 'Body');
    session.createStory('�', 'Replacement');
    session.createStory('😀', 'Pair');
    for (const story of ['body', 'hf:rId1', '�', '\uD800', '\uDC00x', '😀', '']) {
      expect(session.hasStory(story)).toBe(session.storyIds().includes(story));
    }
  } finally {
    session.destroy();
  }
});
