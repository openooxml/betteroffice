import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../../wasm/edit';
import { createYrsSession } from '../../yrs';
import { createYrsSidebarProjection } from './yrsSidebarProjection';

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);

test('a session reuses its projection until the document changes', async () => {
  const session = await createYrsSession({ clientId: 80011 });
  try {
    const [first, second] = session.loadStories([
      { storyId: 'body', paragraphs: [{ text: 'alpha beta' }, { text: 'gamma' }] },
    ]).body;
    const projection = createYrsSidebarProjection(session);
    const before = projection.locToDisplayPoint({ story: 'body', paraId: second!, offset: 1 });
    expect(createYrsSidebarProjection(session)).toBe(projection);

    session.insertText({ story: 'body', paraId: first!, offset: 5 }, ' one', {
      name: 'Ada',
      date: '2026-09-29T00:00:00Z',
    });
    const changed = createYrsSidebarProjection(session);
    expect(changed).not.toBe(projection);
    expect(changed.locToDisplayPoint({ story: 'body', paraId: second!, offset: 1 })).toEqual({
      story: 'body',
      position: before!.position + 4,
    });
  } finally {
    session.destroy();
  }
});

test('a projection reads body segments once for multiple paragraphs', async () => {
  const session = await createYrsSession({ clientId: 80012 });
  try {
    const firstText = 'alpha beta';
    const secondText = 'gamma';
    const [first, second] = session.loadStories([
      { storyId: 'body', paragraphs: [{ text: firstText }, { text: secondText }] },
    ]).body;
    const read = session.storySegments.bind(session);
    const calls: string[] = [];
    session.storySegments = (story) => {
      calls.push(story);
      return read(story);
    };
    const projection = createYrsSidebarProjection(session);
    expect(
      projection.locToDisplayPoint({ story: 'body', paraId: first!, offset: firstText.length })
    ).toEqual({
      story: 'body',
      position: 1 + firstText.length,
    });
    expect(
      projection.locToDisplayPoint({ story: 'body', paraId: second!, offset: secondText.length })
    ).toEqual({
      story: 'body',
      position: firstText.length + 2 + 1 + secondText.length,
    });
    expect(calls.filter((story) => story === 'body')).toHaveLength(1);
  } finally {
    session.destroy();
  }
});
