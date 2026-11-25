import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { EDIT_PARAGRAPH_ID, oneSavePackage } from './__fixtures__/oneSavePackage';
import { createYrsSession, saveYrsDocx } from './index';

describe('seedFromDocx -> applyEdits -> saveYrsDocx', () => {
  const bodyPart = 'word/document.xml';
  let source: ReturnType<typeof unzipContainer>;
  let saved: ReturnType<typeof unzipContainer>;
  let body: string;
  let edited: string;
  let probes: readonly string[];

  beforeAll(async () => {
    await preloadEditWasm(new Uint8Array(readFileSync(
      resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
    )));
    const fixture = oneSavePackage();
    probes = fixture.probes;
    source = unzipContainer(fixture.bytes);
    const session = await createYrsSession();
    try {
      session.seedFromDocx(fixture.bytes);
      const edit = session.applyEdits({
        expectVersion: session.version(),
        steps: [{
          op: 'insertText',
          target: { kind: 'paragraph', story: 'body', paraId: EDIT_PARAGRAPH_ID },
          at: 'end',
          text: ' (edited)',
        }],
      });
      expect(edit).toMatchObject({ ok: true, applied: true });
      saved = unzipContainer((await saveYrsDocx(session)).bytes);
      body = new TextDecoder().decode(saved[bodyPart]);
      edited = body.match(new RegExp(`<w:p\\b[^>]*w14:paraId=["']${EDIT_PARAGRAPH_ID}["'][^>]*>[\\s\\S]*?</w:p>`))?.[0] ?? '';
    } finally {
      session.destroy();
    }
  }, 30_000);

  test('writes the edit and keeps every other part apart from core properties', () => {
    expect([...edited.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]).join(''))
      .toBe('EDIT ME (edited)');
    expect(Object.keys(saved).sort()).toEqual(Object.keys(source).sort());
    for (const [part, bytes] of Object.entries(source)) {
      if (part !== bodyPart && part !== 'docProps/core.xml') expect(saved[part]).toEqual(bytes);
    }
  });

  test.failing('keeps every source probe outside the edited paragraph', () => {
    const outsideEdit = body.replace(edited, '');
    expect(probes.filter((marker) =>
      !outsideEdit.includes(`"${marker}"`) && !outsideEdit.includes(`'${marker}'`)
    )).toEqual([]);
  });

  test.failing('keeps the authors of both move-range starts', () => {
    const authors = Object.entries({ moveFromRangeStart: '10', moveToRangeStart: '20' }).map(([name, id]) => {
      const marker = body.match(new RegExp(`<w:${name}\\b[^>]*w:id=["']${id}["'][^>]*>`))?.[0];
      return marker?.match(/\bw:author=["']([^"']+)["']/)?.[1];
    });
    expect(authors).toEqual(['Author', 'Author']);
  });
});
