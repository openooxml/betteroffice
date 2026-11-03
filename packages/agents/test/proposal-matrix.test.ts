import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { attachDocx, attachXlsx, openDocx, openXlsx } from '../src';
import { fixture } from './fixture';
import { xlsxFixture } from './format-fixtures';

type Mode = 'headless' | 'attached' | 'peer';
type Format = 'docx' | 'xlsx';
interface Hit { target: string; ref: string }
interface ScriptDocument {
  read(): string[];
  grep(cursor?: string): { hit: Hit; cursor: string | null };
  propose(target: string, text: string): Promise<{ id: string }>;
  review(id: string): { status: string; stale: boolean };
  accept(id: string): Promise<unknown>;
  reject(id: string): Promise<unknown>;
  exportRead(id?: string): Promise<string[]>;
  pendingCount(): number;
  userEdit(): void;
}

const exceptions = {
  pendingRegistry: 'Only attached proposals enter the shared engine registry; DOCX editor saves contain their live tracked changes.',
  ownAcceptCursor: 'Attached cursors survive own acceptance; headless cursors retain their previous version expiry.',
  ownAcceptCell: 'Attached cell handles survive own acceptance; headless XLSX handles retain whole-workbook version expiry.',
  unrelatedProposal: 'Attached XLSX proposals use target guards; headless XLSX proposals retain whole-workbook staleness.',
  diagnostics: 'Attached XLSX accept diagnostics come from engine acceptance, including cycleCells 0.',
  acceptDate: 'Attached DOCX refuses accept date; headless tracked acceptance keeps its supplied date.',
  trackedConflict: 'Attached DOCX refuses adjoining tracked revisions with their identity; headless keeps ordinary insertion beside a tracked run.',
  undo: 'Attached DOCX settlement stays outside undo history; headless acceptance remains one undo step.',
};

async function runAgentScript(doc: ScriptDocument, format: Format, mode: Mode) {
  const attached = mode !== 'headless';
  const original = ['draft alpha', 'draft beta', 'draft gamma'];
  expect(doc.read()).toEqual(original);
  const first = doc.grep();
  const second = doc.grep(first.cursor!);
  const accepted = await doc.propose(first.hit.target, 'approved alpha');
  expect(doc.grep(first.cursor!).hit.ref).toBe(second.hit.ref);
  const rejected = await doc.propose(second.hit.target, 'approved beta');
  expect(doc.grep(second.cursor!).hit).toBeDefined();
  expect(doc.pendingCount(), exceptions.pendingRegistry).toBe(attached ? 2 : 0);
  expect(doc.review(accepted.id)).toEqual({ status: 'pending', stale: false });
  expect(doc.review(rejected.id)).toEqual({ status: 'pending', stale: false });
  expect(doc.read()).toEqual(original);
  expect(await doc.exportRead()).toEqual(original);
  expect(await doc.exportRead(accepted.id)).toEqual(['approved alpha', 'draft beta', 'draft gamma']);
  await doc.accept(accepted.id);
  expect(doc.review(accepted.id)).toEqual({ status: 'accepted', stale: false });
  expect(doc.review(rejected.id), exceptions.unrelatedProposal).toEqual({ status: 'pending', stale: format === 'xlsx' && !attached });
  await doc.reject(rejected.id);
  expect(doc.review(rejected.id)).toEqual({ status: 'rejected', stale: false });
  if (attached) {
    expect(doc.grep(second.cursor!).hit, exceptions.ownAcceptCursor).toBeDefined();
    const reused = await doc.propose(second.hit.target, 'review beta');
    await doc.reject(reused.id);
    const acceptedTarget = await doc.propose(first.hit.target, 'review alpha');
    await doc.reject(acceptedTarget.id);
  } else {
    expect(() => doc.grep(second.cursor!), exceptions.ownAcceptCursor).toThrow('edited');
    if (format === 'xlsx') await expect(doc.propose(second.hit.target, 'review beta'), exceptions.ownAcceptCell).rejects.toMatchObject({ code: 'STALE_TARGET' });
    else {
      const reused = await doc.propose(second.hit.target, 'review beta');
      await doc.reject(reused.id);
    }
  }
  expect(await doc.exportRead()).toEqual(['approved alpha', 'draft beta', 'draft gamma']);
  expect(doc.read()).toEqual(['approved alpha', 'draft beta', 'draft gamma']);
  const fresh = doc.grep();
  const stale = await doc.propose(fresh.hit.target, 'review beta');
  doc.userEdit();
  expect(doc.review(stale.id)).toEqual({ status: 'pending', stale: true });
  await expect(doc.accept(stale.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
  await expect(doc.propose(fresh.hit.target, 'review beta')).rejects.toMatchObject({ code: 'STALE_TARGET' });
  expect(() => doc.grep(fresh.cursor!)).toThrow('edited');
  await doc.reject(stale.id);
}

for (const mode of ['headless', 'attached', 'peer'] as const) {
  test(`DOCX shared agent script in ${mode} mode`, async () => {
    const zip = await JSZip.loadAsync(await fixture());
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace(/<w:body>[\s\S]*<w:sectPr>/, '<w:body><w:p><w:r><w:t>draft alpha</w:t></w:r></w:p><w:p><w:r><w:t>draft beta</w:t></w:r></w:p><w:p><w:r><w:t>draft gamma</w:t></w:r></w:p><w:sectPr>'));
    const owner = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
    const doc = mode === 'headless' ? owner : attachDocx(owner.session, mode === 'peer' ? { editPeer: {
      async proposeChanges(request) { await Promise.resolve(); return owner.session.proposeChanges(request); },
      async setProposalStates(request) { await Promise.resolve(); return owner.session.setProposalStates(request); },
      async withdrawProposals(request) { await Promise.resolve(); return owner.session.withdrawProposals(request); },
    } } : {});
    const read = () => doc.list({ story: 'body' }).items.map(p => doc.read(p.ref).text);
    try {
      await runAgentScript({
        read,
        grep(cursor) {
          const page = doc.grep({ query: 'draft', story: 'body', limit: 1, cursor });
          return { hit: { target: page.matches[0].match, ref: page.matches[0].ref }, cursor: page.nextCursor };
        },
        propose: (match, newText) => doc.proposeAsync({ author: 'Agent', edits: [{ match, newText }] }),
        review(id) { const review = doc.review(id); return { status: review.status, stale: review.staleRefs.length > 0 }; },
        accept: id => doc.accept(id),
        reject: id => doc.rejectAsync(id),
        async exportRead(id) {
          const reopened = await openDocx(await doc.export(id));
          try { return reopened.list({ story: 'body' }).items.map(p => reopened.read(p.ref).text); }
          finally { reopened.close(); }
        },
        pendingCount: () => owner.session.getProposals().proposals.length,
        userEdit() { const paragraph = owner.session.paragraphs('body')[1]; owner.session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'User '); },
      }, 'docx', mode);
    } finally { doc.close(); owner.close(); }
  });
}

for (const mode of ['headless', 'attached'] as const) {
  test(`XLSX shared agent script in ${mode} mode`, async () => {
    const owner = await openXlsx(await xlsxFixture());
    for (const [row, text] of ['draft alpha', 'draft beta', 'draft gamma'].entries()) owner.handle.editCell(0, row, 0, text);
    const doc = mode === 'headless' ? owner : attachXlsx(owner.handle);
    try {
      await runAgentScript({
        read: () => doc.readCells({ sheet: 'sheet:0', range: 'A1:A3' }).cells.map(cell => cell.displayText),
        grep(cursor) {
          const page = doc.grep({ query: 'draft', story: 'sheet:0', limit: 1, cursor });
          return { hit: { target: page.matches[0].cell, ref: page.matches[0].ref }, cursor: page.nextCursor };
        },
        async propose(cell, input) { return doc.proposeCells({ author: 'Agent', edits: [{ cell, input }] }); },
        review(id) { const review = doc.review(id); return { status: review.status, stale: review.stale }; },
        async accept(id) {
          const result = await doc.accept(id);
          if (mode === 'attached') expect(result.receipt, exceptions.diagnostics).toMatchObject({ calculation: { cycleCells: 0 } });
          return result;
        },
        async reject(id) { return doc.reject(id); },
        async exportRead(id) {
          const reopened = await openXlsx(await doc.export(id));
          try { return reopened.readCells({ sheet: 'sheet:0', range: 'A1:A3' }).cells.map(cell => cell.displayText); }
          finally { reopened.close(); }
        },
        pendingCount: () => owner.handle.listProposals().length,
        userEdit() { owner.handle.editCell(0, 1, 0, 'User draft beta'); },
      }, 'xlsx', mode);
    } finally { doc.close(); owner.close(); }
  });
}

for (const mode of ['headless', 'attached'] as const) {
  test(`DOCX explicit date, tracked conflict and undo exceptions in ${mode} mode`, async () => {
    const owner = await openDocx(await fixture());
    const doc = mode === 'headless' ? owner : attachDocx(owner.session);
    try {
      const hit = doc.grep({ query: 'Executive summary' }).matches[0];
      const cursor = doc.grep({ query: 'is', limit: 1 }).nextCursor!;
      const proposal = doc.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] });
      const date = '2026-01-01T00:00:00Z';
      if (mode === 'attached') {
        await expect(doc.accept(proposal.id, { tracked: true, date }), exceptions.acceptDate).rejects.toMatchObject({ code: 'ACCEPT_DATE_UNSUPPORTED' });
        expect(doc.review(proposal.id).status).toBe('pending');
        await doc.accept(proposal.id, { tracked: true });
        expect(doc.read(hit.ref).revision).toBe(hit.revision);
        expect(doc.grep({ query: 'is', limit: 1, cursor }).matches).toHaveLength(1);
      } else {
        await doc.accept(proposal.id, { tracked: true, date });
        expect(owner.session.listRevisions().every(revision => revision.date === date), exceptions.acceptDate).toBe(true);
      }
      const paragraph = owner.session.paragraphs('body')[2];
      owner.session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Tracked', { name: 'Reviewer', date });
      const ref = doc.list({ story: 'body' }).items[2];
      const revision = owner.session.listRevisions().find(revision => revision.author === 'Reviewer')!;
      const input = { author: 'Agent', edits: [{ ...ref, start: 7, oldText: '', newText: ' Ordinary ' }] };
      if (mode === 'attached') {
        expect(() => doc.propose(input), exceptions.trackedConflict).toThrow('Reviewer');
        try { doc.propose(input); throw new Error('Expected tracked conflict'); }
        catch (error) { expect(error).toMatchObject({
          code: 'TRACKED_REVISION_CONFLICT',
          conflicts: expect.arrayContaining([expect.objectContaining({ revisionId: revision.revisionId, author: 'Reviewer', kind: 'insertion', range: revision.range, ref: ref.ref })]),
        }); }
        try { doc.propose({ ...input, edits: [{ ...ref, start: 0, oldText: 'Tracked', newText: 'Changed' }] }); throw new Error('Expected tracked conflict'); }
        catch (error) { expect(error).toMatchObject({ code: 'TRACKED_REVISION_CONFLICT' }); }
      } else {
        const insertion = doc.propose(input);
        await doc.accept(insertion.id);
        expect(doc.read(ref.ref).text, exceptions.trackedConflict).toContain('Tracked Ordinary ');
      }
    } finally { doc.close(); owner.close(); }

    const undoOwner = await openDocx(await fixture());
    const undoDoc = mode === 'headless' ? undoOwner : attachDocx(undoOwner.session);
    try {
      const prior = undoOwner.session.paragraphs('body')[1];
      undoOwner.session.addUndoBoundary();
      undoOwner.session.insertText({ story: 'body', paraId: prior.paraId, offset: 0 }, 'Prior ');
      undoOwner.session.addUndoBoundary();
      const hit = undoDoc.grep({ query: 'Executive summary' }).matches[0];
      const proposal = undoDoc.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] });
      await undoDoc.accept(proposal.id);
      const paragraph = undoOwner.session.paragraphs('body')[0];
      undoOwner.session.addUndoBoundary();
      undoOwner.session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Later ');
      undoOwner.session.addUndoBoundary();
      expect(undoOwner.session.undo()).toBe(true);
      expect(undoDoc.read(hit.ref).text).toBe('Updated summary');
      undoOwner.session.undo();
      expect(undoDoc.read(hit.ref).text, exceptions.undo).toBe(mode === 'headless' ? 'Executive summary' : 'Updated summary');
    } finally { undoDoc.close(); undoOwner.close(); }
  });
}
