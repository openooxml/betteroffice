import { afterEach, describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { attachDocx, openDocx, type DocxAgentDocument } from '../src';
import { fixture } from './fixture';

const opened: DocxAgentDocument[] = [];
async function open(extra = 0) {
  const doc = await openDocx(await fixture(extra));
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

describe('grep, read, propose, export', () => {
  test('paginates a large document and returns exact contextual references', async () => {
    const doc = await open(1500);
    let cursor: string | undefined;
    const refs: string[] = [];
    do {
      const page = doc.grep({ query: 'risk assessment', limit: 100, cursor });
      expect(JSON.stringify(page).length).toBeLessThan(45000);
      for (const match of page.matches) {
        const p = doc.read(match.ref);
        expect(p.text.slice(match.start, match.end)).toBe(match.text);
        refs.push(match.ref);
      }
      cursor = page.nextCursor ?? undefined;
    } while (cursor);
    expect(refs.length).toBe(1500);
    expect(new Set(refs).size).toBe(1500);
    expect(doc.list({ headingsOnly: true }).items[0].text).toBe('Executive summary');
    expect(doc.grep({ query: 'Table target' }).matches).toHaveLength(1);
    expect(doc.grep({ query: 'Internal' }).matches).toHaveLength(1);
  });

  test('uses literal case-insensitive matching without regex interpretation', async () => {
    const doc = await open();
    expect(doc.grep({ query: 'project aurora' }).matches.length).toBe(2);
    expect(doc.grep({ query: 'project aurora', caseSensitive: true }).matches).toHaveLength(0);
    expect(doc.grep({ query: '.*' }).matches).toHaveLength(0);
    expect(() => doc.grep({ query: '' })).toThrow('query must');
    expect(() => doc.grep({ query: 'a', limit: -1 })).toThrow('limit must');
  });

  test('edits the selected repeated occurrence through a grep match ID', async () => {
    const doc = await open();
    const hits = doc.grep({ query: 'draft' }).matches;
    const proposal = doc.propose({ author: 'test', edits: [{ match: hits[1].match, newText: 'approved' }] });
    await doc.accept(proposal.id);
    expect(doc.read(hits[0].ref).text).toBe('The forecast is draft; the appendix is approved.');
    expect(() => doc.propose({ author: 'test', edits: [{ match: hits[0].match, newText: 'stale' }] })).toThrow('changed');
    expect(() => doc.propose({ author: 'test', edits: [{ match: 'nonexistent', newText: 'wrong' }] })).toThrow('Run grep again');
  });

  test('preserves offsets across embeds and surrogate pairs', async () => {
    const doc = await open();
    const target = doc.session.paragraphs('body').find(p => p.text.startsWith('Before'))!;
    doc.session.insertPageBreak({ story: 'body', paraId: target.paraId, offset: 6 });
    const [hit] = doc.grep({ query: 'After 😀' }).matches;
    expect(hit.start).toBe(7);
    const p = doc.read(hit.ref);
    expect(p.text).toContain('Before\ufffcAfter 😀');
    expect(doc.grep({ query: 'BeforeAfter' }).matches).toHaveLength(0);
    expect(() => doc.propose({ author: 'test', edits: [{ ...hit, oldText: '\ud83d', newText: 'x', start: 13 }] })).toThrow('Unicode');
    expect(() => doc.propose({ author: 'test', edits: [{ match: hit.match, newText: '\ud83d' }] })).toThrow('Unicode');
    const proposal = doc.propose({ author: 'test', edits: [{ ...hit, oldText: hit.text, newText: 'Updated 😀' }] });
    await doc.accept(proposal.id);
    expect(doc.read(hit.ref).text).toContain('Before\ufffcUpdated 😀');
  });

  test('exports a proposal without mutation and preserves formatting and untouched parts', async () => {
    const doc = await open();
    const [hit] = doc.grep({ query: '€4.2 million' }).matches;
    const before = doc.read(hit.ref);
    const proposal = doc.propose({ author: 'test-agent', note: 'Update revenue', edits: [{ ...hit, oldText: hit.text, newText: '€5.1 million' }] });
    const saved = await doc.export(proposal.id);
    expect(doc.read(hit.ref)).toEqual(before);
    expect(doc.review(proposal.id).status).toBe('pending');
    const zip = await JSZip.loadAsync(saved);
    expect(await zip.file('customXml/preserved.xml')!.async('string')).toBe('<payload>Keep these exact bytes.</payload>');
    expect(await zip.file('word/header1.xml')!.async('string')).toBe(await (await JSZip.loadAsync(await fixture())).file('word/header1.xml')!.async('string'));
    const reopened = await openDocx(saved);
    opened.push(reopened);
    const [changed] = reopened.grep({ query: '€5.1 million' }).matches;
    expect(changed).toBeDefined();
    const run = reopened.read(changed.ref).runs.find(run => run.start <= changed.start && run.end >= changed.end);
    expect(run?.formatting.bold).toBe(true);
    expect(JSON.stringify(run?.formatting)).toContain('0066AA');
    expect(await doc.verify(proposal.id)).toMatchObject({ reopened: true });
  });

  test('rejects ambiguous, stale, overlapping, invalid and protected edits without mutation', async () => {
    const doc = await open();
    const [hit] = doc.grep({ query: 'draft' }).matches;
    const before = doc.read(hit.ref);
    expect(() => doc.propose({ author: 'test', edits: [{ ref: hit.ref, revision: hit.revision, oldText: 'draft', newText: 'final' }] })).toThrow('more than once');
    expect(() => doc.propose({ author: 'test', edits: [{ ...hit, oldText: 'wrong', newText: 'final' }] })).toThrow('does not match');
    expect(() => doc.propose({ author: 'test', edits: [{ ...hit, oldText: 'draft', newText: 'final\nnext' }] })).toThrow('paragraph breaks');
    expect(() => doc.propose({ author: 'test', edits: [0, 1].map(() => ({ ...hit, oldText: 'draft', newText: 'final' })) })).toThrow('non-touching');
    const [embed] = doc.grep({ query: 'Before' }).matches;
    const target = doc.session.paragraphs('body').find(p => p.text.startsWith('Before'))!;
    doc.session.insertPageBreak({ story: 'body', paraId: target.paraId, offset: 6 });
    const updated = doc.read(embed.ref);
    expect(() => doc.propose({ author: 'test', edits: [{ ...updated, start: 0, oldText: 'Before\ufffcAfter', newText: 'Removed' }] })).toThrow('embed');
    expect(doc.read(hit.ref)).toEqual(before);
    const proposal = doc.propose({ author: 'test', edits: [{ ...hit, oldText: 'draft', newText: 'final' }] });
    const p = doc.session.paragraphs(hit.story).find(p => p.text.includes('forecast'))!;
    doc.session.insertText({ story: hit.story, paraId: p.paraId, offset: 0 }, 'Changed ');
    expect(doc.review(proposal.id).staleRefs).toEqual([hit.ref]);
    await expect(doc.accept(proposal.id)).rejects.toThrow('changed');
    expect(() => doc.propose({ author: 'test', edits: [{ ...hit, oldText: 'draft', newText: 'final' }] })).toThrow('changed');
    expect(doc.reject(proposal.id).status).toBe('rejected');
  });

  test('accepts a batch as one undo step, isolated from prior and subsequent user edits', async () => {
    const doc = await open();
    const heading = doc.session.paragraphs('body')[0];
    const loc = { story: 'body', paraId: heading.paraId, offset: 0 };
    doc.session.insertText(loc, 'Prior ');
    const hits = doc.grep({ query: 'draft' }).matches;
    const proposal = doc.propose({ author: 'test', edits: hits.map(hit => ({ ...hit, oldText: hit.text, newText: 'approved' })) });
    await doc.accept(proposal.id);
    doc.session.insertText(loc, 'Later ');
    expect(doc.session.undo()).toBe(true);
    expect(doc.grep({ query: 'Later' }).matches).toHaveLength(0);
    expect(doc.grep({ query: 'approved' }).matches).toHaveLength(3);
    expect(doc.session.undo()).toBe(true);
    expect(doc.grep({ query: 'draft' }).matches).toHaveLength(2);
    expect(doc.grep({ query: 'Prior' }).matches).toHaveLength(1);
    expect(doc.session.redo()).toBe(true);
    expect(doc.grep({ query: 'draft' }).matches).toHaveLength(0);
  });

  test('detects formatting and deletion changes, and permits unrelated edits', async () => {
    const doc = await open(1);
    const [hit] = doc.grep({ query: '€4.2 million' }).matches;
    const proposal = doc.propose({ author: 'test', edits: [{ ...hit, oldText: hit.text, newText: '€5.1 million' }] });
    const tail = doc.session.paragraphs('body').at(-1)!;
    doc.session.insertText({ story: 'body', paraId: tail.paraId, offset: 0 }, 'Peer edit: ');
    await doc.accept(proposal.id);
    expect(doc.grep({ query: '€5.1 million' }).matches).toHaveLength(1);
    expect(doc.grep({ query: 'Peer edit:' }).matches).toHaveLength(1);
    const [newHit] = doc.grep({ query: '€5.1 million' }).matches;
    const next = doc.propose({ author: 'test', edits: [{ ...newHit, oldText: newHit.text, newText: '€6 million' }] });
    const target = doc.session.paragraphs('body').find(p => p.text.includes('€5.1'))!;
    doc.session.formatRange({ story: 'body', start: { paraId: target.paraId, offset: newHit.start }, end: { paraId: target.paraId, offset: newHit.end } }, { italic: true });
    await expect(doc.accept(next.id)).rejects.toThrow('changed');
  });

  test('batch applies against original offsets and supports tracked Word changes', async () => {
    const doc = await open();
    const hits = doc.grep({ query: 'draft' }).matches;
    const proposal = doc.propose({ author: 'GPT 6 Astra', edits: hits.map(hit => ({ ...hit, oldText: hit.text, newText: 'approved' })) });
    await doc.accept(proposal.id, { tracked: true, date: '2026-09-27T12:00:00Z' });
    expect(doc.session.listRevisions().length).toBeGreaterThan(0);
    const zip = await JSZip.loadAsync(await doc.export());
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('GPT 6 Astra');
    expect(xml).toContain('approved');
    expect(xml).toContain('w:del');
  });

  test('invalidates cursors after edits and releases attached sessions without destroying them', async () => {
    const doc = await open(50);
    const cursor = doc.grep({ query: 'Aurora', limit: 1 }).nextCursor!;
    expect(() => doc.grep({ query: 'different', cursor })).toThrow('Repeat grep');
    const p = doc.session.paragraphs('body')[0];
    doc.session.deleteRange({ story: 'body', start: { paraId: p.paraId, offset: 0 }, end: { paraId: p.paraId, offset: 1 } });
    expect(() => doc.grep({ query: 'Aurora', cursor })).toThrow('Repeat grep');
    const attached = attachDocx(doc.session);
    attached.close();
    expect(() => attached.overview()).toThrow('closed');
    expect(doc.overview().paragraphs).toBeGreaterThan(0);
  });

  test('rejects a target edited while asynchronous acceptance begins', async () => {
    const doc = await open();
    const [hit] = doc.grep({ query: '€4.2 million' }).matches;
    const proposal = doc.propose({ author: 'test', edits: [{ match: hit.match, newText: '€5.1 million' }] });
    const acceptance = doc.accept(proposal.id);
    const target = doc.session.paragraphs('body').find(p => p.text.includes('€4.2'))!;
    doc.session.insertText({ story: 'body', paraId: target.paraId, offset: 0 }, 'Concurrent: ');
    await expect(acceptance).rejects.toThrow('changed');
    expect(doc.grep({ query: '€5.1 million' }).matches).toHaveLength(0);
  });

  test('bounds long Unicode reads and discovers direct outline levels', async () => {
    const doc = await open();
    const target = doc.session.createStory('test:long', '😀'.repeat(9000));
    doc.session.setParagraphAttrs({ story: 'test:long', start: { paraId: target.paraId, offset: 0 }, end: { paraId: target.paraId, offset: 0 } }, { other: { outlineLevel: 2 } });
    const item = doc.list({ story: 'test:long' }).items[0];
    expect(item.headingLevel).toBe(3);
    const first = doc.read(item.ref, { length: 1 });
    expect(first.text).toBe('😀');
    expect(first.nextStart).toBe(2);
    const second = doc.read(item.ref, { start: first.nextStart!, length: 16000 });
    expect(second.text.length).toBe(16000);
    expect(second.nextStart).toBe(16002);
    expect(() => doc.read(item.ref, { start: 20000 })).toThrow('start must');
  });
});
