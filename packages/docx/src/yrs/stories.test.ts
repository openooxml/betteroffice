import { afterEach, beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { storiesDocx } from './__fixtures__/storiesDocx';
import { createYrsSession, describeStories, readStorySelection, storyParts, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

const sessions: YrsSession[] = [];
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

async function open(bytes = storiesDocx()) {
  const session = await createYrsSession({ clientId: 76201 });
  sessions.push(session);
  const host = session.openDocx(bytes, true);
  return { session, parts: storyParts(host.document) };
}

function texts(result: ReturnType<YrsSession['readStories']>) {
  if (!result.ok) throw new Error(result.failure.message);
  return Object.fromEntries(
    result.stories.map((story) => [
      story.story,
      'paragraphs' in story ? story.paragraphs.map((paragraph) => paragraph.text) : story.failure.code,
    ])
  );
}

it('describes every story with its kind, container and header or footer uses', async () => {
  const { session, parts } = await open();
  expect(describeStories(session.storyIds(), parts)).toEqual([
    { story: 'body', kind: 'body' },
    { story: 'body:sdt0', kind: 'content-control', parent: 'body' },
    { story: 'body:t0:r0c0', kind: 'table-cell', parent: 'body' },
    { story: 'body:t0:r0c1', kind: 'table-cell', parent: 'body' },
    { story: 'body:t0:r0c1:t0:r0c0', kind: 'table-cell', parent: 'body:t0:r0c1' },
    { story: 'en:1', kind: 'endnote' },
    { story: 'fn:1', kind: 'footnote' },
    {
      story: 'hf:rIdF1',
      kind: 'footer',
      relationshipId: 'rIdF1',
      uses: [
        { sectionIndex: 0, variant: 'default' },
        { sectionIndex: 1, variant: 'default' },
      ],
    },
    { story: 'hf:rIdH1', kind: 'header', relationshipId: 'rIdH1', uses: [{ sectionIndex: 0, variant: 'default' }] },
    { story: 'hf:rIdH1:t0:r0c0', kind: 'table-cell', parent: 'hf:rIdH1' },
    {
      story: 'hf:rIdH2',
      kind: 'header',
      relationshipId: 'rIdH2',
      uses: [
        { sectionIndex: 0, variant: 'first' },
        { sectionIndex: 1, variant: 'first' },
      ],
    },
    { story: 'hf:rIdH3', kind: 'header', relationshipId: 'rIdH3', uses: [{ sectionIndex: 1, variant: 'default' }] },
  ]);
});

it('calls a header or footer story without a host part other', () => {
  expect(describeStories(['hf:rId9', 'hf:rId9:t0:r0c0', 'side'], storyParts(null))).toEqual([
    { story: 'hf:rId9', kind: 'other' },
    { story: 'hf:rId9:t0:r0c0', kind: 'table-cell', parent: 'hf:rId9' },
    { story: 'side', kind: 'other' },
  ]);
});

it('reads every story, the selected kinds or the named stories in one read', async () => {
  const { session, parts } = await open();
  const all = readStorySelection(session, { stories: 'all', view: 'accepted' }, parts);
  expect(all.ok && all.stories.map((story) => story.story)).toEqual(session.storyIds());
  const body = session.readParagraphs({ story: 'body', view: 'accepted' });
  expect(all.ok && all.stories[0]).toEqual({ story: 'body', paragraphs: body.ok ? body.paragraphs : [] });

  expect(texts(readStorySelection(session, { stories: ['header', 'footer'], view: 'accepted' }, parts))).toEqual({
    'hf:rIdF1': ['Footer {{e}}'],
    'hf:rIdH1': ['Header one'],
    'hf:rIdH2': ['First page {{d}}'],
    'hf:rIdH3': ['Header three'],
  });
  expect(
    texts(readStorySelection(session, { stories: ['footnote', 'hf:rIdH1:t0:r0c0', 'missing'], view: 'accepted' }, parts))
  ).toEqual({
    'fn:1': ['Footnote text'],
    'hf:rIdH1:t0:r0c0': ['Header cell {{c}}'],
    missing: 'missing-target',
  });
});

it('refuses a stale version', async () => {
  const { session, parts } = await open();
  const version = session.version();
  expect(readStorySelection(session, { stories: ['body'], view: 'accepted', expectVersion: version }, parts)).toMatchObject({
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
  expect(readStorySelection(session, { stories: 'all', view: 'accepted', expectVersion: version }, parts)).toMatchObject({
    ok: false,
    version: session.version(),
    failure: { code: 'stale-version' },
  });
});

it('reads thousands of stories in one read', async () => {
  const { session, parts } = await open(storiesDocx({ tables: 1500 }));
  const cells = readStorySelection(session, { stories: ['table-cell'], view: 'accepted' }, parts);
  if (!cells.ok) throw new Error(cells.failure.message);
  expect(cells.stories).toHaveLength(3004);
  expect(cells.stories.every((story) => 'paragraphs' in story)).toBe(true);
});
