import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { documentToYrs } from './documentToYrs';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../../../apps/demo/public/betteroffice-demo.docx');
const LEGACY_SEED = resolve(import.meta.dir, '../../../../apps/demo/public/seeds/docx.bin');

function expectEquivalentStories(left: YrsSession, right: YrsSession): void {
  expect(left.storyIds()).toEqual(right.storyIds());
  for (const storyId of left.storyIds()) {
    expect(left.storySegments(storyId)).toEqual(right.storySegments(storyId));
  }
}

describe('DOCX engine seeding parity', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('matches the legacy TypeScript projection structurally', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const parsed = await parseDocx(bytes.buffer);
    const legacy = await createYrsSession({ clientId: 47001 });
    const engine = await createYrsSession({ clientId: 47001 });
    try {
      documentToYrs(legacy, parsed);
      engine.seedFromDocx(bytes);

      expectEquivalentStories(engine, legacy);
      expect(engine.encodeStateVector()).toEqual(legacy.encodeStateVector());

      const legacyStatePeer = await createYrsSession({ clientId: 47002 });
      const engineStatePeer = await createYrsSession({ clientId: 47003 });
      try {
        legacyStatePeer.loadState(legacy.encodeState());
        engineStatePeer.loadState(engine.encodeState());
        const firstParagraph = legacyStatePeer.paragraphs('body')[0];
        legacyStatePeer.insertText(
          { story: 'body', paraId: firstParagraph.paraId, offset: 1 },
          'legacy'
        );
        engineStatePeer.loadState(
          legacyStatePeer.encodeStateAsUpdate(engineStatePeer.encodeStateVector())
        );
        const secondParagraph = engineStatePeer.paragraphs('body')[1];
        engineStatePeer.insertText(
          { story: 'body', paraId: secondParagraph.paraId, offset: 1 },
          'engine'
        );
        legacyStatePeer.loadState(
          engineStatePeer.encodeStateAsUpdate(legacyStatePeer.encodeStateVector())
        );
        expectEquivalentStories(engineStatePeer, legacyStatePeer);
      } finally {
        legacyStatePeer.destroy();
        engineStatePeer.destroy();
      }
    } finally {
      legacy.destroy();
      engine.destroy();
    }
  });

  it('matches the committed pre-change collaboration room structurally', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const legacy = await createYrsSession({ clientId: 47004 });
    const engine = await createYrsSession({ clientId: 1 });
    try {
      legacy.loadState(Uint8Array.from(readFileSync(LEGACY_SEED)));
      engine.seedFromDocx(bytes);

      expectEquivalentStories(engine, legacy);
      expect(engine.encodeStateVector()).toEqual(legacy.encodeStateVector());
    } finally {
      legacy.destroy();
      engine.destroy();
    }
  });

  it('returns thin host metadata and materializes the canonical package on demand', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const parsed = await parseDocx(bytes.buffer);
    const engine = await createYrsSession({ clientId: 47005 });
    const existingRoom = await createYrsSession({ clientId: 47006 });
    try {
      const host = engine.openDocx(bytes, false);

      expect(engine.storyIds()).toEqual([]);
      expect(host.document.package.document.content).toEqual([]);
      expect(
        host.document.package.document.sections?.every((section) => section.content.length === 0)
      ).toBe(true);
      expect(
        [...(host.document.package.headers?.values() ?? [])].every(
          (header) => header.content.length === 0
        )
      ).toBe(true);
      expect(host.document.package.media?.size).toBe(0);
      expect(host.document.package.charts?.size).toBe(0);
      expect(host.referencedFonts.length).toBeGreaterThan(0);

      const materialized = engine.materializeDocx();
      expect(materialized?.package.document.content).toEqual(parsed.package.document.content);
      expect(materialized?.package.document.sections).toEqual(parsed.package.document.sections);
      expect(materialized?.package.headers).toEqual(parsed.package.headers);
      expect(materialized?.package.footers).toEqual(parsed.package.footers);

      existingRoom.loadState(Uint8Array.from(readFileSync(LEGACY_SEED)));
      engine.loadState(existingRoom.encodeState());
      expectEquivalentStories(engine, existingRoom);
    } finally {
      engine.destroy();
      existingRoom.destroy();
    }
  });
});
