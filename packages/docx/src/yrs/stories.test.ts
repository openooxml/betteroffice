import { afterEach, beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { storiesDocx } from './__fixtures__/storiesDocx';
import { createYrsSession, type DocxReadStoriesRequest, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

const sessions: YrsSession[] = [];
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

async function open(bytes = storiesDocx()) {
  const session = await createYrsSession({ clientId: 76201 });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

function texts(session: YrsSession, request: Omit<DocxReadStoriesRequest, 'view'>) {
  const result = session.readStories({ ...request, view: 'accepted' });
  if (!result.ok) throw new Error(result.failure.message);
  return result.stories.map((story) => [
    story.story,
    story.ok ? story.paragraphs.map((paragraph) => paragraph.text) : story.failure.code,
  ]);
}

it('lists every story with its kind, container, root and header or footer uses', async () => {
  const session = await open();
  const listed = session.listStories();
  expect(listed.version).toBe(session.version());
  expect(listed.stories.map(({ part, ...story }) => (part === undefined ? story : { ...story, part: 'part' }))).toEqual([
    { story: 'body', kind: 'body', root: 'body' },
    { story: 'body:sdt0', kind: 'content-control', parent: 'body', root: 'body' },
    { story: 'body:t0:r0c0', kind: 'table-cell', parent: 'body', root: 'body' },
    { story: 'body:t0:r0c1', kind: 'table-cell', parent: 'body', root: 'body' },
    { story: 'body:t0:r0c1:t0:r0c0', kind: 'table-cell', parent: 'body:t0:r0c1', root: 'body' },
    { story: 'en:1', kind: 'endnote', root: 'en:1' },
    { story: 'fn:1', kind: 'footnote', root: 'fn:1' },
    {
      story: 'hf:rIdF1',
      kind: 'footer',
      root: 'hf:rIdF1',
      part: 'part',
      uses: [
        { sectionIndex: 0, variant: 'default' },
        { sectionIndex: 1, variant: 'default' },
      ],
    },
    { story: 'hf:rIdH1', kind: 'header', root: 'hf:rIdH1', part: 'part', uses: [{ sectionIndex: 0, variant: 'default' }] },
    { story: 'hf:rIdH1:t0:r0c0', kind: 'table-cell', parent: 'hf:rIdH1', root: 'hf:rIdH1' },
    {
      story: 'hf:rIdH2',
      kind: 'header',
      root: 'hf:rIdH2',
      part: 'part',
      uses: [
        { sectionIndex: 0, variant: 'first' },
        { sectionIndex: 1, variant: 'first' },
      ],
    },
    { story: 'hf:rIdH3', kind: 'header', root: 'hf:rIdH3', part: 'part', uses: [{ sectionIndex: 1, variant: 'default' }] },
  ]);
  expect(listed.stories.find((story) => story.story === 'hf:rIdH1')?.part).toEndWith('header1.xml');
});

it('reads every story, kinds, roots or ids in one read, keeping the listed order', async () => {
  const session = await open();
  const all = session.readStories({ view: 'accepted' });
  expect(all.ok && all.stories.map((story) => story.story)).toEqual(session.storyIds());
  const body = session.readParagraphs({ story: 'body', view: 'accepted' });
  expect(all.ok && all.stories[0]).toEqual({ story: 'body', ok: true, paragraphs: body.ok ? body.paragraphs : [] });
  expect(all.ok && all.truncated).toBe(false);

  expect(texts(session, { stories: ['footer', 'header'] })).toEqual([
    ['hf:rIdF1', ['Footer {{e}}']],
    ['hf:rIdH1', ['Header one']],
    ['hf:rIdH2', ['First page {{d}}']],
    ['hf:rIdH3', ['Header three']],
  ]);
  expect(texts(session, { stories: ['header'], byRoot: true }).map(([story]) => story)).toEqual([
    'hf:rIdH1',
    'hf:rIdH1:t0:r0c0',
    'hf:rIdH2',
    'hf:rIdH3',
  ]);
  expect(texts(session, { stories: ['missing', 'fn:1', 'footnote', 'hf:rIdH1:t0:r0c0', 'fn:1'] })).toEqual([
    ['missing', 'missing-target'],
    ['fn:1', ['Footnote text']],
    ['hf:rIdH1:t0:r0c0', ['Header cell {{c}}']],
  ]);
});

it('refuses a stale version before reading', async () => {
  const session = await open();
  const version = session.version();
  expect(session.readStories({ stories: ['body'], view: 'accepted', expectVersion: version })).toMatchObject({
    ok: true,
    version,
  });
  const body = session.readParagraphs({ story: 'body', view: 'accepted' });
  const paraId = body.ok ? body.paragraphs[0]!.paraId : '';
  expect(
    session.applyEdits({
      expectVersion: version,
      steps: [{ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId }, at: 'end', text: '!' }],
    }).ok
  ).toBe(true);
  expect(session.readStories({ view: 'accepted', expectVersion: version })).toMatchObject({
    ok: false,
    version: session.version(),
    failure: { code: 'stale-version' },
  });
});

it('reads thousands of stories in one read', async () => {
  const session = await open(storiesDocx({ tables: 1500 }));
  const cells = session.readStories({ stories: ['table-cell'], view: 'accepted' });
  if (!cells.ok) throw new Error(cells.failure.message);
  expect(cells.stories).toHaveLength(3004);
  expect(cells.stories.every((story) => story.ok)).toBe(true);
});

it('lists and reads the stories of a first-page preview', async () => {
  const session = await createYrsSession({ clientId: 76202 });
  sessions.push(session);
  expect(session.openDocxPreview(storiesDocx({ tables: 50 }), 3)).not.toBeNull();
  const listed = session.listStories();
  expect(listed.stories.map((story) => story.story)).toEqual(session.storyIds());
  const read = session.readStories({ view: 'accepted' });
  expect(read.ok && read.stories.map((story) => [story.story, story.ok])).toEqual(
    listed.stories.map((story) => [story.story, true])
  );
});
