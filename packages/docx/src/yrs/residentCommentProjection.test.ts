import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, openCommandEngine, TIMEOUT } from './residentCommandTestSupport';
import { readCommentInsertionRange } from './residentCommentProjection';

afterEach(closeCommandEngines);

const at = (paraId: string, offset: number) => ({ story: 'body', paraId, offset });
const deletion = (id: number) =>
  `<w:del w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>XY</w:delText></w:r></w:del>`;

test('comment insertion ranges refuse search text that occurs more than once', async () => {
  const { engine, resident } = await openCommandEngine(['Seed and Seed', 'Alpha aaa']);
  const [first, second] = engine.paragraphs('body').map(paragraph => paragraph.paraId);
  expect(readCommentInsertionRange(resident, first!, 'Seed')).toEqual({ status: 'refused', reason: 'ambiguous-target' });
  expect(readCommentInsertionRange(resident, second!, 'aa')).toEqual({ status: 'refused', reason: 'ambiguous-target' });
  expect(readCommentInsertionRange(resident, first!, 'and')).toEqual({ story: 'body', start: at(first!, 5), end: at(first!, 8) });
}, TIMEOUT);

test('comment insertion ranges start at the match after a tracked deletion', async () => {
  const { engine, resident } = await openCommandEngine([], { bodyXml:
    `<w:p w14:paraId="00000001"><w:r><w:t>ab</w:t></w:r>${deletion(1)}<w:r><w:t>cd</w:t></w:r></w:p>` +
    `<w:p w14:paraId="00000002">${deletion(2)}<w:r><w:t>ef</w:t></w:r></w:p>` });
  const [first, second] = engine.paragraphs('body').map(paragraph => paragraph.paraId);
  expect(readCommentInsertionRange(resident, first!, 'cd')).toEqual({ story: 'body', start: at(first!, 4), end: at(first!, 6) });
  expect(readCommentInsertionRange(resident, second!, 'ef')).toEqual({ story: 'body', start: at(second!, 2), end: at(second!, 4) });
}, TIMEOUT);
