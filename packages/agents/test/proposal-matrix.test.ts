import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { attachDocx, attachXlsx, openDocx, openXlsx } from '../src';
import { editFixture, fixture } from './fixture';
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

for (const mode of ['headless', 'attached', 'peer'] as const) {
  if (mode !== 'headless') for (const action of ['accept', 'reject'] as const) {
    test(`DOCX partial editor decision followed by agent ${action} in ${mode} mode`, async () => {
      const zip = await JSZip.loadAsync(await fixture());
      const xml = await zip.file('word/document.xml')!.async('string');
      zip.file('word/document.xml', xml.replace('Executive summary', 'alpha beta gamma'));
      const owner = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
      const doc = attachDocx(owner.session, mode === 'peer' ? { editPeer: {
        async proposeChanges(request) { return owner.session.proposeChanges(request); },
        async setProposalStates(request) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
      } } : {});
      try {
        const hits = ['alpha', 'beta', 'gamma'].map(query => doc.grep({ query }).matches[0]);
        const input = { author: 'Agent', edits: hits.map((hit, index) => ({ match: hit.match, newText: ['ALPHA', 'BETA', 'GAMMA'][index] })) };
        const proposal = mode === 'peer' ? await doc.proposeAsync(input) : doc.propose(input);
        const snapshot = owner.session.getProposals();
        const middle = snapshot.proposals.find(record => record.id.endsWith(':1'))!;
        expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
          changes: [{ id: middle.id, state: action === 'accept' ? 'rejected' : 'accepted' }],
        }).ok).toBe(true);
        expect(doc.review(proposal.id).status).toBe('pending');
        if (action === 'accept') await doc.accept(proposal.id);
        else if (mode === 'peer') await doc.rejectAsync(proposal.id);
        else doc.reject(proposal.id);
        expect(doc.read(hits[0].ref).text).toBe(action === 'accept' ? 'ALPHA beta GAMMA' : 'alpha BETA gamma');
        expect(doc.review(proposal.id).status).toBe('accepted');
        expect(owner.session.getProposals().proposals).toHaveLength(0);
      } finally { doc.close(); owner.close(); }
    });
  }

  for (const start of [0, 3]) {
    test(`DOCX saved match after insertion at boundary ${start} in ${mode} mode`, async () => {
      const zip = await JSZip.loadAsync(await fixture());
      const xml = await zip.file('word/document.xml')!.async('string');
      zip.file('word/document.xml', xml.replace('Executive summary', 'abcdef'));
      const owner = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
      const doc = mode === 'headless' ? owner : attachDocx(owner.session, mode === 'peer' ? { editPeer: {
        async proposeChanges(request) { return owner.session.proposeChanges(request); },
        async setProposalStates(request) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
      } } : {});
      try {
        const hit = doc.grep({ query: 'abc' }).matches[0];
        const input = { author: 'Agent', edits: [{ ref: hit.ref, revision: hit.revision, start, oldText: '', newText: 'XY' }] };
        const proposal = mode === 'peer' ? await doc.proposeAsync(input) : doc.propose(input);
        await doc.accept(proposal.id);
        expect(doc.read(hit.ref).text).toBe(start === 0 ? 'XYabcdef' : 'abcXYdef');
        const reuse = { author: 'Agent', edits: [{ match: hit.match, newText: 'Z' }] };
        if (mode === 'headless') expect(() => doc.propose(reuse)).toThrow('Target text or formatting changed');
        else {
          const replacement = mode === 'peer' ? await doc.proposeAsync(reuse) : doc.propose(reuse);
          expect(replacement.changes[0]).toMatchObject({ start: start === 0 ? 2 : 0, oldText: 'abc' });
          await doc.accept(replacement.id);
          expect(doc.read(hit.ref).text).toBe(start === 0 ? 'XYZdef' : 'ZXYdef');
          const reopened = await openDocx(await doc.export());
          try { expect(reopened.read(reopened.grep({ query: 'Z' }).matches[0].ref).text).toBe(start === 0 ? 'XYZdef' : 'ZXYdef'); }
          finally { reopened.close(); }
        }
      } finally { doc.close(); owner.close(); }
    });
  }

  for (const sample of [
    { text: 'aaab', start: 1, oldText: 'aab', newText: 'ccb', expected: 'accb' },
    { text: 'ababa', start: 2, oldText: 'aba', newText: 'xyz', expected: 'abxyz' },
    { text: 'zaaab', start: 0, oldText: 'zaaab', newText: 'zaccb', expected: 'zaccb' },
    { text: '😀aaab', start: 3, oldText: 'aab', newText: 'ccb', expected: '😀accb' },
  ]) {
    test(`DOCX overlapping replacement at ${sample.start} in ${JSON.stringify(sample.text)} in ${mode} mode`, async () => {
      const zip = await JSZip.loadAsync(await fixture());
      const xml = await zip.file('word/document.xml')!.async('string');
      zip.file('word/document.xml', xml.replace('Executive summary', sample.text));
      const owner = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
      const doc = mode === 'headless' ? owner : attachDocx(owner.session, mode === 'peer' ? { editPeer: {
        async proposeChanges(request) { return owner.session.proposeChanges(request); },
        async setProposalStates(request) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
      } } : {});
      try {
        const hit = doc.grep({ query: sample.oldText }).matches[0];
        const edit = hit.start === sample.start ? { match: hit.match, newText: sample.newText }
          : { ref: hit.ref, revision: hit.revision, start: sample.start, oldText: sample.oldText, newText: sample.newText };
        const input = { author: 'Agent', edits: [edit] };
        const proposal = mode === 'peer' ? await doc.proposeAsync(input) : doc.propose(input);
        expect(doc.read(hit.ref).text).toBe(sample.text);
        expect(doc.review(proposal.id).staleRefs).toEqual([]);
        const preview = await openDocx(await doc.export(proposal.id));
        try { expect(preview.read(preview.list({ story: 'body' }).items[0].ref).text).toBe(sample.expected); }
        finally { preview.close(); }
        await doc.accept(proposal.id);
        expect(doc.read(hit.ref).text).toBe(sample.expected);
        expect(owner.session.getProposals().proposals).toHaveLength(0);
      } finally { doc.close(); owner.close(); }
    });
  }

  for (const newText of ['Updated note.', 'Closing note!.', 'Closing note. More']) {
    test(`DOCX replacement after a table for ${JSON.stringify(newText)} in ${mode} mode`, async () => {
      const owner = await openDocx(await editFixture());
      const doc = mode === 'headless' ? owner : attachDocx(owner.session, mode === 'peer' ? { editPeer: {
        async proposeChanges(request) { return owner.session.proposeChanges(request); },
        async setProposalStates(request) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
      } } : {});
      try {
        const hit = doc.grep({ query: 'Closing note.' }).matches[0];
        const input = { author: 'Agent', edits: [{ match: hit.match, newText }] };
        const proposal = mode === 'peer' ? await doc.proposeAsync(input) : doc.propose(input);
        expect(doc.read(hit.ref).text).toBe('\ufffcClosing note.');
        await doc.accept(proposal.id);
        expect(doc.read(hit.ref).text).toBe(`\ufffc${newText}`);
        const reopened = await openDocx(await doc.export());
        try { expect(reopened.grep({ query: newText }).matches).toHaveLength(1); }
        finally { reopened.close(); }
      } finally { doc.close(); owner.close(); }
    });
  }
}

for (const mode of ['headless', 'attached'] as const) {
  for (const mutation of ['dependent', 'style'] as const) {
    test(`XLSX ${mutation} changes stale review and accept in ${mode} mode`, async () => {
      const owner = await openXlsx(await xlsxFixture());
      for (const [col, input] of ['120', '135', '150', '=SUM(B2:D2)'].entries()) owner.handle.editCell(0, 1, col + 1, input);
      const doc = mode === 'headless' ? owner : attachXlsx(owner.handle);
      try {
        const cell = doc.readCells({ sheet: 'sheet:0', range: 'B2' }).cells[0].cell;
        const proposal = doc.proposeCells({ author: 'Agent', edits: [{ cell, input: '200' }] });
        expect(doc.review(proposal.id).stale).toBe(false);
        if (mutation === 'dependent') owner.handle.editCell(0, 1, 2, '140');
        else owner.handle.patchRangeStyle(0, 'B2', { bold: true });
        const version = owner.handle.version();
        const pending = owner.handle.listProposals();
        expect(doc.review(proposal.id)).toMatchObject({ status: 'pending', stale: true });
        expect(doc.listProposals()[0].stale).toBe(true);
        expect(owner.handle.listProposals()).toEqual(pending);
        await expect(doc.export(proposal.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
        await expect(doc.accept(proposal.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
        expect(owner.handle.version()).toBe(version);
        expect(owner.handle.cell(0, 1, 1).input).toBe('120');
        expect(doc.readCells({ sheet: 'sheet:0', range: 'E2' }).cells[0].value).toEqual({ kind: 'number', value: mutation === 'dependent' ? 410 : 405 });
        expect(doc.review(proposal.id).status).toBe('pending');
        expect(owner.handle.listProposals()).toHaveLength(mode === 'attached' ? 1 : 0);
        doc.reject(proposal.id);
      } finally { doc.close(); owner.close(); }
    });
  }
}

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
        async propose(match, newText) {
          const input = { author: 'Agent', edits: [{ match, newText: newText.split(' ')[0] }] };
          return mode === 'peer' ? doc.proposeAsync(input) : doc.propose(input);
        },
        review(id) { const review = doc.review(id); return { status: review.status, stale: review.staleRefs.length > 0 }; },
        accept: id => doc.accept(id),
        async reject(id) { return mode === 'peer' ? doc.rejectAsync(id) : doc.reject(id); },
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
