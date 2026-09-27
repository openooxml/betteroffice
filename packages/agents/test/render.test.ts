import { expect, test } from 'bun:test';
import { openDocx } from '../src';
import { renderDocxPage } from '../src/render';
import { fixture } from './fixture';

test('renders actual before/after pages without changing the live document', async () => {
  const doc = await openDocx(await fixture(), { renderer: renderDocxPage });
  try {
    const [hit] = doc.grep({ query: 'Executive summary' }).matches;
    const proposal = doc.propose({ author: 'test', edits: [{ ...hit, oldText: hit.text, newText: 'Annual report 2026' }] });
    const before = await doc.render(1);
    const after = await doc.render(1, proposal.id);
    expect([...before.png.slice(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(before.pageCount).toBeGreaterThan(0);
    expect(before.png).not.toEqual(after.png);
    expect(doc.read(hit.ref).text).toBe('Executive summary');
    await expect(doc.render(1000)).rejects.toThrow('Choose a page');
  } finally { doc.close(); }
});
