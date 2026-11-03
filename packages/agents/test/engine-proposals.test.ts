import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { attachDocx, attachXlsx, DocumentToolError, openDocx, openXlsx, ProposalDecidedInEditorError } from '../src';
import { editFixture, fixture } from './fixture';
import { xlsxFixture } from './format-fixtures';

test('attached XLSX proposals belong to the engine and reflect editor decisions', async () => {
  const owner = await openXlsx(await xlsxFixture());
  const agent = attachXlsx(owner.handle);
  try {
    const propose = (input: string) => agent.proposeCells({ author: 'Agent', note: 'Review cell', edits: [{ cell: agent.readCells({ sheet: 'sheet:0', range: 'B3' }).cells[0].cell, input }] });
    const accepted = propose('123');
    const [engine] = owner.handle.listProposals();
    expect(engine).toMatchObject({ agentId: 'Agent', note: 'Review cell', cells: [{ a1: 'B3', newText: '123' }] });
    expect(agent.readCells({ sheet: 'sheet:0', range: 'B3' }).cells[0].value).toEqual({ kind: 'number', value: 100 });
    owner.handle.acceptProposal(engine!.id);
    owner.handle.editCell(0, 2, 1, '456');
    expect(agent.review(accepted.id).status).toBe('accepted');
    expect(agent.listProposals()[0].status).toBe('accepted');
    const rejected = propose('789');
    owner.handle.rejectProposal(owner.handle.listProposals()[0]!.id);
    expect(agent.review(rejected.id).status).toBe('rejected');
    const own = propose('=C3*2');
    await agent.accept(own.id);
    expect(owner.handle.listProposals()).toHaveLength(0);
    expect(owner.handle.cell(0, 2, 1).input).toBe('=C3*2');
    const stale = propose('999');
    owner.handle.editCell(0, 2, 1, '321');
    await expect(agent.accept(stale.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
    expect(owner.handle.cell(0, 2, 1).input).toBe('321');
    agent.reject(stale.id);
    expect(owner.handle.listProposals()).toHaveLength(0);
  } finally { agent.close(); owner.close(); }
});

test('XLSX proposal acceptance permits unrelated user edits', async () => {
  const owner = await openXlsx(await xlsxFixture());
  const book = attachXlsx(owner.handle);
  try {
    const cell = book.readCells({ sheet: 'sheet:0', range: 'B3' }).cells[0].cell;
    const proposal = book.proposeCells({ author: 'Agent', edits: [{ cell, input: '123' }] });
    book.handle.editCell(0, 100, 0, 'Unrelated');
    expect(book.review(proposal.id).stale).toBe(false);
    await book.accept(proposal.id);
    expect(book.handle.cell(0, 2, 1).input).toBe('123');
  } finally { book.close(); owner.close(); }
});

test('attached XLSX string stale refusals never apply direct edits', async () => {
  const owner = await openXlsx(await xlsxFixture());
  let accepts = 0;
  let writes = 0;
  const handle = new Proxy(owner.handle, { get(target, key) {
    if (key === 'acceptProposal') return () => { accepts++; throw new Error('Stale proposal: B3'); };
    if (key === 'applyEdits' || key === 'editCell') return () => { writes++; throw new Error('Unexpected direct edit'); };
    return Reflect.get(target, key);
  } });
  const doc = attachXlsx(handle);
  try {
    const cell = doc.readCells({ sheet: 'sheet:0', range: 'B3' }).cells[0].cell;
    const proposal = doc.proposeCells({ author: 'Agent', edits: [{ cell, input: '200' }] });
    const pending = owner.handle.listProposals();
    const version = owner.handle.version();
    await expect(doc.accept(proposal.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
    expect(accepts).toBe(1);
    expect(writes).toBe(0);
    expect(owner.handle.version()).toBe(version);
    expect(owner.handle.cell(0, 2, 1).input).toBe('100');
    expect(owner.handle.listProposals()).toEqual(pending);
    expect(doc.review(proposal.id).status).toBe('pending');
  } finally { doc.close(); owner.close(); }
});

for (const peer of [false, true]) {
  for (const newText of ['Updated note.', 'Closing note!.', 'Closing note. More']) {
    test(`attached DOCX maps leading block embeds for ${JSON.stringify(newText)} with peer=${peer}`, async () => {
      const owner = await openDocx(await editFixture());
      const paragraph = owner.session.paragraphs('body')[0];
      const start = owner.session.locateParagraph('body', paragraph.paraId).start;
      owner.session.createStory('body:offset-control', 'control');
      owner.session.applyRawOps('body', [
        { op: 'insertEmbed', index: start, kind: 'pageBreak' },
        { op: 'insertEmbed', index: start + 1, kind: 'columnBreak' },
        { op: 'insertEmbed', index: start + 2, kind: 'blockSdt', payload: { story: 'body:offset-control' } },
      ]);
      const doc = attachDocx(owner.session, peer ? { editPeer: {
        async proposeChanges(request) { return owner.session.proposeChanges(request); },
        async setProposalStates(request) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
      } } : {});
      try {
        for (const query of ['Quarterly report', 'Closing note.']) {
          const hit = doc.grep({ query }).matches[0];
          expect(hit.start).toBe(query === 'Quarterly report' ? 3 : 1);
          const replacement = query === 'Quarterly report' ? `Updated ${query}` : newText;
          const input = { author: 'Agent', edits: [{ match: hit.match, newText: replacement }] };
          const proposal = peer ? await doc.proposeAsync(input) : doc.propose(input);
          expect(doc.review(proposal.id).staleRefs).toEqual([]);
          const read = owner.session.readParagraphs({ story: 'body', view: 'accepted' });
          expect(read.ok && read.paragraphs.some(p => p.text === replacement)).toBe(true);
          await doc.accept(proposal.id);
          expect(doc.read(hit.ref).text).toBe(`${'\ufffc'.repeat(hit.start)}${replacement}`);
        }
        expect(owner.session.getProposals().proposals).toHaveLength(0);
      } finally { doc.close(); owner.close(); }
    });
  }
}

test('attached DOCX proposals show engine revisions while pending reads and exports retain original text', async () => {
  const owner = await openDocx(await fixture());
  const agent = attachDocx(owner.session);
  try {
    const [hit] = agent.grep({ query: 'Executive summary' }).matches;
    const before = agent.read(hit.ref);
    const proposal = agent.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] });
    expect(owner.session.getProposals().proposals).toHaveLength(1);
    expect(owner.session.listRevisions().every(revision => revision.author === 'Agent')).toBe(true);
    const editorSave = await JSZip.loadAsync(await owner.export());
    expect(await editorSave.file('word/document.xml')!.async('string')).toMatch(/<w:(ins|del)\b/);
    expect(agent.read(hit.ref)).toEqual(before);
    expect(agent.grep({ query: 'Updated' }).matches).toHaveLength(0);
    expect(agent.review(proposal.id).staleRefs).toEqual([]);
    const pending = await JSZip.loadAsync(await agent.export());
    expect(await pending.file('word/document.xml')!.async('string')).toContain('Executive');
    expect(await pending.file('word/document.xml')!.async('string')).not.toMatch(/<w:(ins|del)\b/);
    const preview = await openDocx(await agent.export(proposal.id));
    try { expect(preview.grep({ query: 'Updated summary' }).matches).toHaveLength(1); }
    finally { preview.close(); }
    agent.reject(proposal.id);
    expect(owner.session.getProposals().proposals).toHaveLength(0);
    expect(owner.session.listRevisions()).toHaveLength(0);
  } finally { agent.close(); owner.close(); }
});

test('attached DOCX keeps unidentified engine protection refusals out of revision conflict errors', async () => {
  const owner = await openDocx(await fixture());
  const session = new Proxy(owner.session, { get(target, key) {
    if (key === 'proposeChanges') return () => ({ ok: false, version: target.version(), failure: { code: 'tracked-revision-conflict', message: 'Range is protected' } });
    return Reflect.get(target, key);
  } });
  const doc = attachDocx(session);
  try {
    const hit = doc.grep({ query: 'Executive summary' }).matches[0];
    expect(owner.session.listRevisions()).toHaveLength(0);
    let refusal: unknown;
    try { doc.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] }); }
    catch (error) { refusal = error; }
    expect(refusal).toMatchObject({ code: 'PROTECTED_CONTENT', message: 'Range is protected' });
    expect(doc.listProposals()).toHaveLength(0);
    expect(owner.session.getProposals().proposals).toHaveLength(0);
    expect(doc.read(hit.ref).text).toBe('Executive summary');
  } finally { doc.close(); owner.close(); }
});

test('attached DOCX sends the known trimmed offset for overlapping replacement text', async () => {
  for (const sample of [
    { text: 'aaab', oldText: 'aab', start: 1, newText: 'ccb', search: 'aa', replaceWith: 'cc', expected: 'accb' },
    { text: 'ababa', oldText: 'aba', start: 2, newText: 'xyz', search: 'aba', replaceWith: 'xyz', expected: 'abxyz' },
  ]) {
    const zip = await JSZip.loadAsync(await fixture());
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace('Executive summary', sample.text));
    const owner = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
    const requests: Array<Parameters<typeof owner.session.proposeChanges>[0]> = [];
    const doc = attachDocx(owner.session, { editPeer: {
      async proposeChanges(request) { requests.push(request); return owner.session.proposeChanges(request); },
      async setProposalStates(request) { return owner.session.setProposalStates(request); },
      async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
    } });
    try {
      const hit = doc.grep({ query: sample.oldText }).matches[0];
      const proposal = await doc.proposeAsync({ author: 'Agent', edits: [{ ref: hit.ref, revision: hit.revision, start: sample.start, oldText: sample.oldText, newText: sample.newText }] });
      expect(requests[0].proposals[0]).toMatchObject({ op: 'replaceText', at: { offset: sample.start }, search: sample.search, replaceWith: sample.replaceWith });
      expect(doc.read(hit.ref).text).toBe(sample.text);
      expect(owner.session.getProposals().proposals).toHaveLength(1);
      await doc.accept(proposal.id);
      expect(doc.read(hit.ref).text).toBe(sample.expected);
      expect(owner.session.getProposals().proposals).toHaveLength(0);
    } finally { doc.close(); owner.close(); }
  }
});

for (const peer of [false, true]) {
  for (const editorState of ['accepted', 'rejected'] as const) for (const action of ['accept', 'tracked', 'reject', 'discard'] as const) {
    test(`attached DOCX refuses ${action} after editor ${editorState}, peer=${peer}`, async () => {
      const zip = await JSZip.loadAsync(await fixture());
      const xml = await zip.file('word/document.xml')!.async('string');
      zip.file('word/document.xml', xml.replace('Executive summary', 'alpha beta gamma'));
      const owner = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
      const calls: string[] = [];
      const session = new Proxy(owner.session, { get(target, key) {
        if (key === 'setProposalStates') return (request: Parameters<typeof target.setProposalStates>[0]) => {
          calls.push('decide');
          return target.setProposalStates(request);
        };
        if (key === 'withdrawProposals') return (request: Parameters<typeof target.withdrawProposals>[0]) => {
          calls.push('withdraw');
          return target.withdrawProposals(request);
        };
        return Reflect.get(target, key);
      } });
      const doc = attachDocx(session, peer ? { editPeer: {
        async proposeChanges(request) { await Promise.resolve(); return session.proposeChanges(request); },
        async setProposalStates(request) { await Promise.resolve(); return session.setProposalStates(request); },
        async withdrawProposals(request) { await Promise.resolve(); return session.withdrawProposals(request); },
      } } : {});
      try {
        const hits = ['alpha', 'beta', 'gamma'].map(query => doc.grep({ query }).matches[0]);
        const input = { author: 'Agent', edits: hits.map((hit, index) => ({ match: hit.match, newText: ['ALPHA', 'BETA', 'GAMMA'][index] })) };
        const proposal = peer ? await doc.proposeAsync(input) : doc.propose(input);
        const snapshot = owner.session.getProposals();
        const middle = snapshot.proposals.find(record => record.id.endsWith(':1'))!;
        expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
          changes: [{ id: middle.id, state: editorState }],
        }).ok).toBe(true);
        const records = owner.session.getProposals().proposals.map(({ id, state }) => ({ id, state }));
        expect(doc.review(proposal.id)).toMatchObject({ status: 'pending', records });
        expect(doc.listProposals().find(value => value.id === proposal.id)).toMatchObject({ status: 'pending', records });
        const before = owner.session.encodeState();
        const registry = owner.session.getProposals();
        const revisions = owner.session.listRevisions();
        const refuse = async () => {
          if (action === 'accept' || action === 'tracked') return doc.accept(proposal.id, { tracked: action === 'tracked' });
          if (action === 'reject') return peer ? doc.rejectAsync(proposal.id) : doc.reject(proposal.id);
          return peer ? doc.discardAsync(proposal.id) : doc.discard(proposal.id);
        };
        const error = await refuse().catch(error => error);
        expect(error).toBeInstanceOf(DocumentToolError);
        expect(error).toBeInstanceOf(ProposalDecidedInEditorError);
        expect(error).toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR', details: { records: [{ id: middle.id, state: editorState }] } });
        expect(calls).toEqual([]);
        expect(owner.session.encodeState()).toEqual(before);
        expect(owner.session.getProposals()).toEqual(registry);
        expect(owner.session.listRevisions()).toEqual(revisions);
        expect(doc.review(proposal.id)).toMatchObject({ status: 'pending', records });
        if (peer) for (const mutation of [() => doc.reject(proposal.id), () => doc.discard(proposal.id)]) {
          expect(mutation).toThrow(ProposalDecidedInEditorError);
        }
        expect(owner.session.setProposalStates({ expectVersion: registry.version, expectPreviewVersion: registry.previewVersion,
          changes: registry.proposals.filter(record => record.state === 'proposed').map(record => ({ id: record.id, state: 'accepted' as const })),
        }).ok).toBe(true);
        expect(doc.review(proposal.id).status).toBe('accepted');
        expect(owner.session.withdrawProposals({ expectVersion: owner.session.version(), ids: registry.proposals.map(record => record.id) }).ok).toBe(true);
        expect(owner.session.getProposals().proposals).toHaveLength(0);
        expect(doc.read(hits[0].ref).text).toBe(`ALPHA ${editorState === 'accepted' ? 'BETA' : 'beta'} GAMMA`);
        expect(doc.listProposals().find(value => value.id === proposal.id)?.status).toBe('accepted');
        const finished = owner.session.encodeState();
        await expect(refuse()).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR' });
        expect(owner.session.encodeState()).toEqual(finished);
      } finally { doc.close(); owner.close(); }
    });
  }

  for (const editorState of ['accepted', 'rejected'] as const) {
    test(`attached DOCX editor ${editorState} expires only touched paragraph matches, peer=${peer}`, async () => {
      const owner = await openDocx(await fixture());
      const doc = attachDocx(owner.session, peer ? { editPeer: {
        async proposeChanges(request) { return owner.session.proposeChanges(request); },
        async setProposalStates(request) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
      } } : {});
      try {
        const hits = doc.grep({ query: 'draft' }).matches;
        const untouched = doc.grep({ query: 'Executive summary' }).matches[0];
        const page = doc.grep({ query: 'draft', limit: 1 });
        const input = { author: 'Agent', edits: [{ match: hits[0].match, newText: 'confirmed' }] };
        const proposal = peer ? await doc.proposeAsync(input) : doc.propose(input);
        const snapshot = owner.session.getProposals();
        expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
          changes: [{ id: snapshot.proposals[0].id, state: editorState }],
        }).ok).toBe(true);
        expect(doc.read(hits[1].ref).revision).toBeGreaterThan(hits[1].revision);
        expect(doc.read(untouched.ref).revision).toBe(untouched.revision);
        expect(() => doc.grep({ query: 'draft', limit: 1, cursor: page.nextCursor! })).toThrow('edited');
        const reuse = { author: 'Agent', edits: [{ match: hits[1].match, newText: 'final' }] };
        await expect(doc.proposeAsync(reuse)).rejects.toMatchObject({ code: 'STALE_TARGET' });
        await expect(doc.proposeAsync({ author: 'Agent', edits: [{ ref: hits[1].ref, revision: hits[1].revision, start: hits[1].start, oldText: 'draft', newText: 'final' }] })).rejects.toMatchObject({ code: 'STALE_TARGET' });
        const unaffected = { author: 'Agent', edits: [{ match: untouched.match, newText: 'Overview' }] };
        const next = peer ? await doc.proposeAsync(unaffected) : doc.propose(unaffected);
        await doc.accept(next.id);
        expect(doc.read(untouched.ref)).toMatchObject({ revision: untouched.revision, text: 'Overview' });
        const fresh = doc.grep({ query: 'draft' }).matches.at(-1)!;
        const refreshed = await doc.proposeAsync({ author: 'Agent', edits: [{ match: fresh.match, newText: 'final' }] });
        await doc.rejectAsync(refreshed.id);
        expect(doc.review(proposal.id).status).toBe(editorState);
      } finally { doc.close(); owner.close(); }
    });

    test(`attached DOCX retains editor ownership after decisions reverse, peer=${peer}, state=${editorState}`, async () => {
      const owner = await openDocx(await fixture());
      const doc = attachDocx(owner.session, peer ? { editPeer: {
        async proposeChanges(request) { return owner.session.proposeChanges(request); },
        async setProposalStates(request) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
      } } : {});
      try {
        const hit = doc.grep({ query: 'draft' }).matches[0];
        const proposal = await doc.proposeAsync({ author: 'Agent', edits: [{ match: hit.match, newText: 'confirmed' }] });
        for (const state of [editorState, 'proposed'] as const) {
          const snapshot = owner.session.getProposals();
          expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
            changes: [{ id: snapshot.proposals[0].id, state }],
          }).ok).toBe(true);
        }
        expect(doc.review(proposal.id).status).toBe('pending');
        const before = owner.session.encodeState();
        await expect(doc.accept(proposal.id)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR' });
        await expect(doc.discardAsync(proposal.id)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR' });
        expect(owner.session.encodeState()).toEqual(before);
      } finally { doc.close(); owner.close(); }
    });
  }

  test(`attached DOCX refuses an editor withdrawal outside agent decisions, peer=${peer}`, async () => {
    const owner = await openDocx(await fixture());
    const doc = attachDocx(owner.session, peer ? { editPeer: {
      async proposeChanges(request) { return owner.session.proposeChanges(request); },
      async setProposalStates(request) { return owner.session.setProposalStates(request); },
      async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
    } } : {});
    try {
      const hit = doc.grep({ query: 'draft' }).matches[0];
      const proposal = await doc.proposeAsync({ author: 'Agent', edits: [{ match: hit.match, newText: 'confirmed' }] });
      const snapshot = owner.session.getProposals();
      const record = snapshot.proposals[0];
      expect(owner.session.withdrawProposals({ expectVersion: snapshot.version, ids: [record.id] }).ok).toBe(true);
      const before = owner.session.encodeState();
      await expect(doc.rejectAsync(proposal.id)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR', details: { records: [{ id: record.id, state: 'rejected' }] } });
      expect(owner.session.encodeState()).toEqual(before);
      expect(doc.review(proposal.id)).toMatchObject({ status: 'rejected', records: [{ id: record.id, state: 'rejected' }] });
    } finally { doc.close(); owner.close(); }
  });

  test(`attached DOCX refuses eviction of an editor decided group, peer=${peer}`, async () => {
    const owner = await openDocx(await fixture());
    const doc = attachDocx(owner.session, peer ? { editPeer: {
      async proposeChanges(request) { return owner.session.proposeChanges(request); },
      async setProposalStates(request) { return owner.session.setProposalStates(request); },
      async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
    } } : {});
    try {
      const hit = doc.grep({ query: 'draft' }).matches[0];
      const input = { author: 'Agent', edits: [{ match: hit.match, newText: 'confirmed' }] };
      const first = await doc.proposeAsync(input);
      const snapshot = owner.session.getProposals();
      const record = snapshot.proposals[0];
      expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
        changes: [{ id: record.id, state: 'rejected' }],
      }).ok).toBe(true);
      expect(owner.session.withdrawProposals({ expectVersion: owner.session.version(), ids: [record.id] }).ok).toBe(true);
      const fresh = doc.grep({ query: 'draft' }).matches[0];
      const repeat = { author: 'Agent', edits: [{ match: fresh.match, newText: 'confirmed' }] };
      for (let index = 1; index < 64; index++) {
        const proposal = await doc.proposeAsync(repeat);
        await doc.rejectAsync(proposal.id);
      }
      const before = owner.session.encodeState();
      await expect(doc.proposeAsync(repeat)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR', details: { records: [{ id: record.id, state: 'rejected' }] } });
      expect(owner.session.encodeState()).toEqual(before);
      expect(doc.listProposals()).toHaveLength(64);
      expect(doc.review(first.id).status).toBe('rejected');
    } finally { doc.close(); owner.close(); }
  });

  for (const sample of [
    { query: 'abc', start: 3, oldText: '', newText: 'XY', text: 'abcXYdef', matchStart: 0, matchText: 'abc', expected: 'ZXYdef' },
    { query: 'abc', start: 0, oldText: '', newText: 'XY', text: 'XYabcdef', matchStart: 2, matchText: 'abc', expected: 'XYZdef' },
    { query: 'abc', start: 1, oldText: 'b', newText: 'XY', text: 'aXYcdef', matchStart: 0, matchText: 'aXYc', expected: 'Zdef' },
    { query: 'abc', start: 2, oldText: 'cd', newText: 'XY', text: 'abXYef', matchStart: 0, matchText: 'abXY', expected: 'Zef' },
  ]) {
    test(`attached DOCX saved match after ${JSON.stringify(sample.oldText)} at ${sample.start} with peer=${peer}`, async () => {
      const zip = await JSZip.loadAsync(await fixture());
      const xml = await zip.file('word/document.xml')!.async('string');
      zip.file('word/document.xml', xml.replace('Executive summary', 'abcdef'));
      const owner = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
      const doc = attachDocx(owner.session, peer ? { editPeer: {
        async proposeChanges(request) { return owner.session.proposeChanges(request); },
        async setProposalStates(request) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
      } } : {});
      try {
        const hit = doc.grep({ query: sample.query }).matches[0];
        const input = { author: 'Agent', edits: [{ ref: hit.ref, revision: hit.revision, start: sample.start, oldText: sample.oldText, newText: sample.newText }] };
        const proposal = peer ? await doc.proposeAsync(input) : doc.propose(input);
        await doc.accept(proposal.id);
        expect(doc.read(hit.ref)).toMatchObject({ revision: hit.revision, text: sample.text });
        const reuse = { author: 'Agent', edits: [{ match: hit.match, newText: 'Z' }] };
        const replacement = peer ? await doc.proposeAsync(reuse) : doc.propose(reuse);
        expect(replacement.changes[0]).toMatchObject({ start: sample.matchStart, oldText: sample.matchText, newText: 'Z' });
        await doc.accept(replacement.id);
        expect(doc.read(hit.ref).text).toBe(sample.expected);
      } finally { doc.close(); owner.close(); }
    });
  }
}

test('attached DOCX recognizes an editor decision while an async agent decision is waiting', async () => {
  const owner = await openDocx(await fixture());
  let editorState: Uint8Array | undefined;
  const doc = attachDocx(owner.session, { editPeer: {
    async proposeChanges(request) { return owner.session.proposeChanges(request); },
    async setProposalStates(request) {
      await Promise.resolve();
      expect(owner.session.setProposalStates(request).ok).toBe(true);
      editorState = owner.session.encodeState();
      return owner.session.setProposalStates(request);
    },
    async withdrawProposals() { throw new Error('Agent must not withdraw an editor decided group.'); },
  } });
  try {
    const hit = doc.grep({ query: 'draft' }).matches[0];
    const proposal = await doc.proposeAsync({ author: 'Agent', edits: [{ match: hit.match, newText: 'confirmed' }] });
    const record = owner.session.getProposals().proposals[0];
    await expect(doc.accept(proposal.id)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR', details: { records: [{ id: record.id, state: 'accepted' }] } });
    if (!editorState) throw new Error('Editor state was not captured.');
    expect(owner.session.encodeState()).toEqual(editorState);
    expect(owner.session.getProposals().proposals).toHaveLength(1);
    await expect(doc.proposeAsync({ author: 'Agent', edits: [{ match: hit.match, newText: 'final' }] })).rejects.toMatchObject({ code: 'STALE_TARGET' });
  } finally { doc.close(); owner.close(); }
});

for (const state of ['accepted', 'rejected'] as const) {
  test(`DOCX agent observes editor ${state} preview and withdrawal`, async () => {
    const owner = await openDocx(await fixture());
    const agent = attachDocx(owner.session);
    try {
      const [hit] = agent.grep({ query: 'Executive summary' }).matches;
      const proposal = agent.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] });
      const snapshot = owner.session.getProposals();
      const ids = snapshot.proposals.map(record => record.id);
      expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion, changes: ids.map(id => ({ id, state })) }).ok).toBe(true);
      expect(agent.listProposals()[0].status).toBe(state);
      expect(agent.review(proposal.id).status).toBe(state);
      expect(agent.grep({ query: state === 'accepted' ? 'Updated summary' : 'Executive summary' }).matches).toHaveLength(1);
      expect(owner.session.withdrawProposals({ expectVersion: owner.session.version(), ids }).ok).toBe(true);
      expect(agent.review(proposal.id).status).toBe(state);
    } finally { agent.close(); owner.close(); }
  });
}

for (const tracked of [false, true]) {
  test(`DOCX agent acceptance uses engine revisions with tracked=${tracked}`, async () => {
    const owner = await openDocx(await fixture());
    const doc = attachDocx(owner.session);
    try {
      const [hit] = doc.grep({ query: 'Executive summary' }).matches;
      const proposal = doc.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] });
      const revisionIds = doc.session.listRevisions().map(revision => revision.revisionId);
      const dates = doc.session.listRevisions().map(revision => revision.date);
      expect(revisionIds.length).toBeGreaterThan(0);
      await expect(doc.accept(proposal.id, { tracked, date: '2026-01-01T00:00:00Z' })).rejects.toMatchObject({ code: 'ACCEPT_DATE_UNSUPPORTED' });
      expect(doc.review(proposal.id).status).toBe('pending');
      expect(await doc.accept(proposal.id, { tracked })).toMatchObject({ id: proposal.id, status: 'accepted', tracked, changes: 1 });
      expect(doc.session.getProposals().proposals).toHaveLength(0);
      expect(doc.session.listRevisions().map(revision => revision.revisionId)).toEqual(tracked ? revisionIds : []);
      if (tracked) expect(doc.session.listRevisions().map(revision => revision.date)).toEqual(dates);
      const zip = await JSZip.loadAsync(await doc.export());
      const xml = await zip.file('word/document.xml')!.async('string');
      expect(xml).toContain('Updated');
      expect(/<w:ins\b/.test(xml)).toBe(tracked);
      expect(doc.review(proposal.id).status).toBe('accepted');
    } finally { doc.close(); owner.close(); }
  });
}

test('attached XLSX edit peer replays agent proposals and decisions to the editor worker', async () => {
  const { readFile } = await import('node:fs/promises');
  const { createInProcessPair } = await import('../../../shared/office-session/testing/inProcessTransport');
  const { createWorkbookSession, hydratePeer } = await import('../../xlsx/src/session/client');
  const { createWorkbookSessionHost } = await import('../../xlsx/src/session/host');
  const { createWorkbookEditPeer } = await import('../../xlsx/src/session/editPeer');
  const { workbookSessionInternals } = await import('../../xlsx/src/session/replay');
  const pair = createInProcessPair();
  createWorkbookSessionHost(pair.host);
  const wasm = new Uint8Array(await readFile(new URL('../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url))).buffer;
  const session = await createWorkbookSession(await xlsxFixture(), { wasm, retainPeerHydration: true }, pair.client);
  const peer = await hydratePeer(session);
  const internal = workbookSessionInternals.get(session)!;
  const replay = internal.replay;
  const replies: unknown[] = [];
  internal.replay = async envelope => {
    const reply = await replay(envelope);
    if (envelope.op.method === 'propose') replies.push(reply.result);
    return reply;
  };
  const edits = createWorkbookEditPeer({ session, peer });
  const agent = attachXlsx(peer, { editPeer: edits });
  try {
    const propose = (input: string) => agent.proposeCells({ author: 'Agent', edits: [{ cell: agent.readCells({ sheet: 'sheet:0', range: 'B3' }).cells[0].cell, input }] });
    const proposal = propose('345');
    await edits.flush();
    const [engine] = peer.listProposals();
    expect(replies[0]).toEqual(engine);
    edits.acceptProposal(engine!.id);
    await edits.flush();
    expect(agent.review(proposal.id).status).toBe('accepted');
    const own = propose('678');
    await agent.accept(own.id);
    await edits.flush();
    const read = await session.call.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3' } }] });
    expect(read.ok && read.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 678 });
    const rejected = propose('999');
    agent.reject(rejected.id);
    await edits.flush();
    expect(peer.listProposals()).toHaveLength(0);
    const stale = propose('111');
    edits.editCell(0, 2, 1, '222');
    await expect(agent.accept(stale.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
    await edits.flush();
    expect(edits.state).toBe('ready');
    agent.reject(stale.id);
    for (const mutation of ['dependent', 'style'] as const) {
      for (const [col, input] of ['120', '135', '150', '=SUM(B2:D2)'].entries()) edits.editCell(0, 1, col + 1, input);
      const cell = agent.readCells({ sheet: 'sheet:0', range: 'B2' }).cells[0].cell;
      const proposal = agent.proposeCells({ author: 'Agent', edits: [{ cell, input: '200' }] });
      if (mutation === 'dependent') edits.editCell(0, 1, 2, '140');
      else edits.patchRangeStyle(0, 'B2', { bold: true });
      await edits.flush();
      const pending = peer.listProposals();
      const version = peer.version();
      expect(agent.review(proposal.id).stale).toBe(true);
      expect(agent.listProposals().find(p => p.id === proposal.id)).toMatchObject({ status: 'pending', stale: true });
      expect(peer.listProposals()).toEqual(pending);
      await expect(agent.accept(proposal.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
      await edits.flush();
      expect(edits.state).toBe('ready');
      expect(peer.version()).toBe(version);
      expect(peer.cell(0, 1, 1).input).toBe('120');
      const remote = await session.call.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B2:E2' } }] });
      expect(remote.ok && remote.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 120 });
      expect(remote.ok && remote.ranges[0].cells[0][3].value).toEqual({ kind: 'number', value: mutation === 'dependent' ? 410 : 405 });
      expect(peer.listProposals()).toEqual(pending);
      expect(agent.review(proposal.id).status).toBe('pending');
      agent.reject(proposal.id);
    }
    await edits.flush();
    let refusedAccepts = 0;
    let directWrites = 0;
    const guarded = new Proxy(peer, { get(target, key) {
      if (key === 'applyEdits' || key === 'editCell') return () => { directWrites++; throw new Error('Unexpected direct edit'); };
      return Reflect.get(target, key);
    } });
    const refusing = new Proxy(edits, { get(target, key) {
      if (key === 'acceptProposal') return () => { refusedAccepts++; throw new Error('Stale proposal: B3'); };
      return Reflect.get(target, key);
    } });
    const refusingAgent = attachXlsx(guarded, { editPeer: refusing });
    try {
      const cell = refusingAgent.readCells({ sheet: 'sheet:0', range: 'B3' }).cells[0].cell;
      const proposal = refusingAgent.proposeCells({ author: 'Agent', edits: [{ cell, input: '333' }] });
      await edits.flush();
      const pending = peer.listProposals();
      const sequence = edits.sentSequence;
      await expect(refusingAgent.accept(proposal.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
      await edits.flush();
      expect(refusedAccepts).toBe(1);
      expect(directWrites).toBe(0);
      expect(edits.sentSequence).toBe(sequence);
      expect(peer.listProposals()).toEqual(pending);
      expect(refusingAgent.review(proposal.id).status).toBe('pending');
      const remote = await session.call.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3' } }] });
      expect(remote.ok && remote.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 222 });
      refusingAgent.reject(proposal.id);
      await edits.flush();
    } finally { refusingAgent.close(); }
  } finally { agent.close(); edits.dispose(); peer.dispose(); await session.dispose(); }
}, 30000);

test('DOCX proposals on separate ranges preserve original agent offsets', async () => {
  const owner = await openDocx(await fixture());
  const doc = attachDocx(owner.session);
  try {
    const hits = doc.grep({ query: 'draft' }).matches;
    const first = doc.propose({ author: 'Agent', edits: [{ match: hits[0].match, newText: 'confirmed' }] });
    const second = doc.propose({ author: 'Agent', edits: [{ match: hits[1].match, newText: 'final' }] });
    expect(doc.review(first.id).staleRefs).toEqual([]);
    expect(doc.review(second.id).staleRefs).toEqual([]);
    const preview = await openDocx(await doc.export(second.id));
    try { expect(preview.read(preview.grep({ query: 'forecast' }).matches[0].ref).text).toBe('The forecast is draft; the appendix is final.'); }
    finally { preview.close(); }
    doc.reject(first.id);
    await doc.accept(second.id);
    expect(doc.read(hits[0].ref).text).toBe('The forecast is draft; the appendix is final.');
  } finally { doc.close(); owner.close(); }
});

test('DOCX pending render uses a private original view and selected preview', async () => {
  const rendered: string[] = [];
  const owner = await openDocx(await fixture());
  const doc = attachDocx(owner.session, { renderer: async session => {
    rendered.push(session.paragraphs('body')[0].text);
    return { png: new Uint8Array(), page: 1, pageCount: 1, width: 1, height: 1, warnings: [] };
  } });
  try {
    const [hit] = doc.grep({ query: 'Executive summary' }).matches;
    const proposal = doc.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] });
    await doc.render();
    await doc.render(1, proposal.id);
    expect(rendered).toEqual(['Executive summary', 'Updated summary']);
    expect(doc.review(proposal.id).status).toBe('pending');
    expect(doc.session.getProposals().proposals).toHaveLength(1);
  } finally { doc.close(); owner.close(); }
});

test('XLSX proposal notifications refresh a live subscriber without document edits', async () => {
  const owner = await openXlsx(await xlsxFixture());
  const book = attachXlsx(owner.handle);
  const lengths: number[] = [];
  const off = book.handle.onProposalChange(proposals => { lengths.push(proposals.length); });
  try {
    const version = book.handle.version();
    const cell = book.readCells({ sheet: 'sheet:0', range: 'B3' }).cells[0].cell;
    const proposal = book.proposeCells({ author: 'Agent', edits: [{ cell, input: '123' }] });
    const [engine] = book.handle.listProposals();
    expect(book.handle.getProposalStatus(engine!.id)).toBe('pending');
    expect(book.handle.version()).toBe(version);
    expect(lengths).toEqual([1]);
    book.reject(proposal.id);
    expect(book.handle.getProposalStatus(engine!.id)).toBe('rejected');
    expect(lengths).toEqual([1, 0]);
    off();
    book.proposeCells({ author: 'Agent', edits: [{ cell, input: '456' }] });
    expect(lengths).toEqual([1, 0]);
    expect(book.handle.getProposalStatus('missing')).toBeUndefined();
  } finally { off(); book.close(); owner.close(); }
});

test('DOCX async attachment routes proposal mutations through the editor peer', async () => {
  const owner = await openDocx(await fixture());
  const calls: string[] = [];
  const session = new Proxy(owner.session, { get(target, key) {
    if (key === 'proposeChanges' || key === 'setProposalStates' || key === 'withdrawProposals') {
      return () => { throw new Error('Synchronous proposal registry is mirrored'); };
    }
    return Reflect.get(target, key);
  } });
  const agent = attachDocx(session, { editPeer: {
    async proposeChanges(request) { calls.push('propose'); await Promise.resolve(); return owner.session.proposeChanges(request); },
    async setProposalStates(request) { calls.push('decide'); await Promise.resolve(); return owner.session.setProposalStates(request); },
    async withdrawProposals(request) { calls.push(request.keepRevisions ? 'keep' : 'settle'); await Promise.resolve(); return owner.session.withdrawProposals(request); },
  } });
  try {
    const propose = async (query: string, newText: string) => {
      const [hit] = agent.grep({ query }).matches;
      return agent.proposeAsync({ author: 'Agent', edits: [{ match: hit.match, newText }] });
    };
    const rejected = await propose('Executive summary', 'Updated summary');
    expect(owner.session.getProposals().proposals).toHaveLength(1);
    expect(agent.grep({ query: 'Executive summary' }).matches).toHaveLength(1);
    expect(await agent.rejectAsync(rejected.id)).toEqual({ id: rejected.id, status: 'rejected' });
    const accepted = await propose('Executive summary', 'Revised summary');
    await agent.accept(accepted.id);
    expect(owner.session.listRevisions()).toHaveLength(0);
    expect(agent.grep({ query: 'Revised summary' }).matches).toHaveLength(1);
    const tracked = await propose('€4.2 million', '€5.1 million');
    await agent.accept(tracked.id, { tracked: true });
    expect(owner.session.listRevisions().length).toBeGreaterThan(0);
    const discarded = await propose('Table target', 'Table review');
    await agent.discardAsync(discarded.id);
    expect(owner.session.getProposals().proposals).toHaveLength(0);
    expect(() => agent.review(discarded.id)).toThrow('Proposal missing');
    expect(calls).toEqual(['propose', 'decide', 'settle', 'propose', 'decide', 'settle', 'propose', 'decide', 'keep', 'propose', 'decide', 'settle']);
  } finally { agent.close(); owner.close(); }
});

for (const peer of [false, true]) {
  test(`attached DOCX keeps refs, matches, cursors and pending ranges through own decisions with peer=${peer}`, async () => {
    const owner = await openDocx(await fixture());
    const doc = attachDocx(owner.session, peer ? { editPeer: {
      async proposeChanges(request) { await Promise.resolve(); return owner.session.proposeChanges(request); },
      async setProposalStates(request) { await Promise.resolve(); return owner.session.setProposalStates(request); },
      async withdrawProposals(request) { await Promise.resolve(); return owner.session.withdrawProposals(request); },
    } } : {});
    try {
      const hits = doc.grep({ query: 'draft' }).matches;
      const page = doc.grep({ query: 'draft', limit: 1 });
      const first = await doc.proposeAsync({ author: 'Agent', edits: [{ match: hits[0].match, newText: 'confirmed' }] });
      const second = await doc.proposeAsync({ author: 'Agent', edits: [{ match: hits[1].match, newText: 'final' }] });
      await doc.accept(first.id);
      expect(doc.read(hits[0].ref).revision).toBe(hits[0].revision);
      expect(doc.review(second.id).staleRefs).toEqual([]);
      expect(doc.grep({ query: 'draft', limit: 1, cursor: page.nextCursor! }).matches[0].start).toBe(hits[1].start + 4);
      await doc.rejectAsync(second.id);
      const reused = await doc.proposeAsync({ author: 'Agent', edits: [{ match: hits[1].match, newText: 'final' }] });
      await doc.accept(reused.id);
      expect(doc.read(hits[0].ref).text).toBe('The forecast is confirmed; the appendix is final.');
      const acceptedTarget = await doc.proposeAsync({ author: 'Agent', edits: [{ match: hits[0].match, newText: 'approved' }] });
      await doc.rejectAsync(acceptedTarget.id);
      const reference = doc.read(hits[0].ref);
      const exact = await doc.proposeAsync({ author: 'Agent', edits: [{ ...reference, revision: hits[0].revision, start: 0, oldText: 'The', newText: 'Our' }] });
      await doc.rejectAsync(exact.id);
      const cursor = doc.grep({ query: 'is', limit: 1 }).nextCursor!;
      const paragraph = owner.session.paragraphs('body')[2];
      owner.session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'User ');
      expect(() => doc.grep({ query: 'is', limit: 1, cursor })).toThrow('edited');
      await expect(doc.proposeAsync({ author: 'Agent', edits: [{ ...reference, start: 0, oldText: 'The', newText: 'Our' }] })).rejects.toMatchObject({ code: 'STALE_TARGET' });
    } finally { doc.close(); owner.close(); }
  });
}

test('an interleaved user edit during an async DOCX proposal does not revive refs or cursors', async () => {
  const owner = await openDocx(await fixture());
  const agent = attachDocx(owner.session, { editPeer: {
    async proposeChanges(request) {
      const result = owner.session.proposeChanges(request);
      const paragraph = owner.session.paragraphs('body')[2];
      owner.session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'User ');
      return result;
    },
    async setProposalStates(request) { return owner.session.setProposalStates(request); },
    async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
  } });
  try {
    const hits = agent.grep({ query: 'draft' }).matches;
    const cursor = agent.grep({ query: 'draft', limit: 1 }).nextCursor!;
    const proposal = await agent.proposeAsync({ author: 'Agent', edits: [{ match: hits[0].match, newText: 'confirmed' }] });
    expect(() => agent.grep({ query: 'draft', limit: 1, cursor })).toThrow('edited');
    await expect(agent.proposeAsync({ author: 'Agent', edits: [{ match: hits[1].match, newText: 'final' }] })).rejects.toMatchObject({ code: 'STALE_TARGET' });
    expect(agent.review(proposal.id).staleRefs).toEqual([hits[0].ref]);
  } finally { agent.close(); owner.close(); }
});

async function proposalRaceFixture() {
  const zip = await JSZip.loadAsync(await fixture());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('Executive summary', 'alpha beta gamma gamma'));
  return openDocx(await zip.generateAsync({ type: 'uint8array' }));
}

for (const peer of [false, true]) {
  for (const state of ['accepted', 'rejected'] as const) {
    test(`attached DOCX selected previews refuse editor ${state} first, peer=${peer}`, async () => {
      const owner = await proposalRaceFixture();
      let rendered = false;
      const doc = attachDocx(owner.session, { renderer: async () => { rendered = true; return { png: new Uint8Array(), page: 1, pageCount: 1, width: 1, height: 1, warnings: [] }; }, ...(peer ? { editPeer: {
        async proposeChanges(request: Parameters<typeof owner.session.proposeChanges>[0]) { return owner.session.proposeChanges(request); },
        async setProposalStates(request: Parameters<typeof owner.session.setProposalStates>[0]) { return owner.session.setProposalStates(request); },
        async withdrawProposals(request: Parameters<typeof owner.session.withdrawProposals>[0]) { return owner.session.withdrawProposals(request); },
      } } : {}) });
      try {
        const hits = ['alpha', 'beta'].map(query => doc.grep({ query }).matches[0]);
        const input = { author: 'Agent', edits: hits.map((hit, i) => ({ match: hit.match, newText: ['ALPHA', 'BETA'][i] })) };
        const proposal = peer ? await doc.proposeAsync(input) : doc.propose(input);
        const snapshot = owner.session.getProposals();
        const record = snapshot.proposals[0];
        expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion, changes: [{ id: record.id, state }] }).ok).toBe(true);
        const before = owner.session.encodeState();
        for (const preview of [() => doc.render(1, proposal.id), () => doc.export(proposal.id), () => doc.verify(proposal.id)]) {
          await expect(preview()).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR', details: { records: [{ id: record.id, state }] } });
        }
        expect(rendered).toBe(false);
        expect(owner.session.encodeState()).toEqual(before);
        expect(owner.session.getProposals().proposals[1].state).toBe('proposed');
      } finally { doc.close(); owner.close(); }
    });
  }

  for (const editorState of ['accepted', 'rejected'] as const) for (const finish of ['decided', 'reverse', 'withdraw'] as const) test(`attached DOCX observes editor ${editorState} before proposal registration, peer=${peer}, finish=${finish}`, async () => {
    const owner = await proposalRaceFixture();
    const doc = attachDocx(owner.session, peer ? { editPeer: {
      async proposeChanges(request) { return owner.session.proposeChanges(request); },
      async setProposalStates(request) { return owner.session.setProposalStates(request); },
      async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
    } } : {});
    let decided = false;
    const unsubscribe = owner.session.onProposalChange(snapshot => {
      if (decided || !snapshot.proposals.length) return;
      decided = true;
      expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
        changes: [{ id: snapshot.proposals[0].id, state: editorState }],
      }).ok).toBe(true);
      if (finish === 'reverse') {
        const current = owner.session.getProposals();
        expect(owner.session.setProposalStates({ expectVersion: current.version, expectPreviewVersion: current.previewVersion,
          changes: [{ id: snapshot.proposals[0].id, state: 'proposed' }],
        }).ok).toBe(true);
      } else if (finish === 'withdraw') expect(owner.session.withdrawProposals({ expectVersion: owner.session.version(), ids: [snapshot.proposals[0].id] }).ok).toBe(true);
    });
    try {
      const alpha = doc.grep({ query: 'alpha' }).matches[0];
      const beta = doc.grep({ query: 'beta' }).matches[0];
      const cursor = doc.grep({ query: 'gamma', limit: 1 }).nextCursor!;
      const input = { author: 'Agent', edits: [{ match: alpha.match, newText: 'ALPHA' }] };
      const proposal = peer ? await doc.proposeAsync(input) : doc.propose(input);
      expect(doc.review(proposal.id).status).toBe(finish === 'reverse' ? 'pending' : editorState);
      await expect(doc.rejectAsync(proposal.id)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR', details: { records: [{ state: editorState }] } });
      await expect(doc.proposeAsync({ author: 'Agent', edits: [{ match: beta.match, newText: 'BETA' }] })).rejects.toMatchObject({ code: 'STALE_TARGET' });
      expect(() => doc.grep({ query: 'gamma', limit: 1, cursor })).toThrow('edited');
    } finally { unsubscribe(); doc.close(); owner.close(); }
  });
}

for (const phase of ['propose', 'propose-after', 'set', 'set-after', 'withdraw', 'withdraw-after'] as const) {
  test(`attached DOCX expires refs across an editor decision during async ${phase}`, async () => {
    const owner = await proposalRaceFixture();
    let armed = false;
    const decideEditor = () => {
      if (!armed) return;
      armed = false;
      const snapshot = owner.session.getProposals();
      expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
        changes: [{ id: snapshot.proposals[0].id, state: 'rejected' }],
      }).ok).toBe(true);
    };
    const doc = attachDocx(owner.session, { editPeer: {
      async proposeChanges(request) {
        await Promise.resolve();
        if (phase === 'propose') decideEditor();
        const result = owner.session.proposeChanges(request);
        if (phase === 'propose-after') decideEditor();
        return result;
      },
      async setProposalStates(request) {
        await Promise.resolve();
        if (phase === 'set') decideEditor();
        const result = owner.session.setProposalStates(request);
        if (phase === 'set-after') decideEditor();
        return result;
      },
      async withdrawProposals(request) {
        await Promise.resolve();
        if (phase === 'withdraw') decideEditor();
        const result = owner.session.withdrawProposals(request);
        if (phase === 'withdraw-after') decideEditor();
        return result;
      },
    } });
    try {
      const [alpha, beta, gamma] = ['alpha', 'beta', 'gamma'].map(query => doc.grep({ query }).matches[0]);
      const cursor = doc.grep({ query: 'gamma', limit: 1 }).nextCursor!;
      await doc.proposeAsync({ author: 'Agent', edits: [{ match: alpha.match, newText: 'ALPHA' }] });
      const input = { author: 'Agent', edits: [{ match: beta.match, newText: 'BETA' }] };
      if (phase === 'propose' || phase === 'propose-after') {
        armed = true;
        await expect(doc.proposeAsync(input)).rejects.toMatchObject({ code: 'STALE_TARGET' });
      } else {
        const second = await doc.proposeAsync(input);
        armed = true;
        await expect(doc.accept(second.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
        if (phase === 'withdraw-after') expect(owner.session.getProposals().proposals).toHaveLength(1);
        else expect(owner.session.getProposals().proposals[1].state).toBe('proposed');
      }
      expect(doc.read(gamma.ref).revision).toBeGreaterThan(gamma.revision);
      await expect(doc.proposeAsync({ author: 'Agent', edits: [{ match: gamma.match, newText: 'GAMMA' }] })).rejects.toMatchObject({ code: 'STALE_TARGET' });
      expect(() => doc.grep({ query: 'gamma', limit: 1, cursor })).toThrow('edited');
    } finally { doc.close(); owner.close(); }
  });
}

for (const phase of ['set', 'set-after', 'withdraw', 'withdraw-one'] as const) for (const action of ['accept', 'tracked', 'reject', 'discard'] as const) {
  test(`attached DOCX rolls back only its own records after editor races ${action} during ${phase}`, async () => {
    const owner = await proposalRaceFixture();
    let armed = false;
    let withdrawals = 0;
    const decideEditor = () => {
      if (!armed) return;
      armed = false;
      const snapshot = owner.session.getProposals();
      expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
        changes: [{ id: snapshot.proposals[0].id, state: (action === 'accept' || action === 'tracked') ? 'rejected' : 'accepted' }],
      }).ok).toBe(true);
      if (phase === 'withdraw-one') expect(owner.session.withdrawProposals({ expectVersion: owner.session.version(), ids: [snapshot.proposals[0].id] }).ok).toBe(true);
    };
    const doc = attachDocx(owner.session, { editPeer: {
      async proposeChanges(request) { return owner.session.proposeChanges(request); },
      async setProposalStates(request) {
        await Promise.resolve();
        if (phase === 'set') decideEditor();
        const result = owner.session.setProposalStates(request);
        if (phase === 'set-after') decideEditor();
        return result;
      },
      async withdrawProposals(request) { withdrawals++; await Promise.resolve(); if (phase === 'withdraw' || phase === 'withdraw-one') decideEditor(); return owner.session.withdrawProposals(request); },
    } });
    try {
      const hits = ['alpha', 'beta'].map(query => doc.grep({ query }).matches[0]);
      const proposal = await doc.proposeAsync({ author: 'Agent', edits: hits.map((hit, i) => ({ match: hit.match, newText: ['ALPHA', 'BETA'][i] })) });
      armed = true;
      const result = (action === 'accept' || action === 'tracked') ? doc.accept(proposal.id, { tracked: action === 'tracked' }) : action === 'reject' ? doc.rejectAsync(proposal.id) : doc.discardAsync(proposal.id);
      await expect(result).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR' });
      expect(withdrawals).toBe(phase === 'withdraw' || phase === 'withdraw-one' ? 1 : 0);
      const states: ('proposed' | 'accepted' | 'rejected')[] = [(action === 'accept' || action === 'tracked') ? 'rejected' : 'accepted', 'proposed'];
      const remaining: typeof states = phase === 'withdraw-one' ? ['proposed'] : states;
      expect(owner.session.getProposals().proposals.map(record => record.state)).toEqual(remaining);
      expect(doc.review(proposal.id).records?.map(record => record.state)).toEqual(states);
      await expect(doc.accept(proposal.id)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR' });
    } finally { doc.close(); owner.close(); }
  });
}

for (const finish of ['reverse', 'withdraw'] as const) {
  test(`attached DOCX retains the actual editor state after matching in-flight decision and ${finish}`, async () => {
    const owner = await proposalRaceFixture();
    let armed = false;
    const doc = attachDocx(owner.session, { editPeer: {
      async proposeChanges(request) { return owner.session.proposeChanges(request); },
      async setProposalStates(request) {
        await Promise.resolve();
        if (armed) {
          armed = false;
          expect(owner.session.setProposalStates(request).ok).toBe(true);
          const current = owner.session.getProposals();
          if (finish === 'reverse') expect(owner.session.setProposalStates({ expectVersion: current.version, expectPreviewVersion: current.previewVersion,
            changes: request.changes.map(({ id }) => ({ id, state: 'proposed' })),
          }).ok).toBe(true);
          else expect(owner.session.withdrawProposals({ expectVersion: current.version, ids: request.changes.map(({ id }) => id) }).ok).toBe(true);
        }
        return owner.session.setProposalStates(request);
      },
      async withdrawProposals() { throw new Error('Agent withdrawal must be refused.'); },
    } });
    try {
      const hit = doc.grep({ query: 'alpha' }).matches[0];
      const proposal = await doc.proposeAsync({ author: 'Agent', edits: [{ match: hit.match, newText: 'ALPHA' }] });
      const id = owner.session.getProposals().proposals[0].id;
      armed = true;
      await expect(doc.accept(proposal.id)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR', details: { records: [{ id, state: 'accepted' }] } });
      await expect(doc.rejectAsync(proposal.id)).rejects.toMatchObject({ code: 'PROPOSAL_DECIDED_IN_EDITOR', details: { records: [{ id, state: 'accepted' }] } });
    } finally { doc.close(); owner.close(); }
  });
}

for (const scenario of ['lagging', 'rollback'] as const) {
  test(`attached DOCX does not infer editor ownership from ${scenario} snapshots`, async () => {
    const owner = await proposalRaceFixture();
    let lagging: ReturnType<typeof owner.session.getProposals> | undefined;
    let replay = false;
    let probe = '';
    let fail = scenario === 'rollback';
    const session = new Proxy(owner.session, { get(target, key) {
      if (key === 'getProposals') return () => replay && lagging ? lagging : target.getProposals();
      return Reflect.get(target, key);
    } });
    const doc = attachDocx(session, { editPeer: {
      async proposeChanges(request) { return owner.session.proposeChanges(request); },
      async setProposalStates(request) {
        lagging = owner.session.getProposals();
        const result = owner.session.setProposalStates(request);
        replay = scenario === 'lagging';
        return result;
      },
      async withdrawProposals(request) {
        if (fail) { fail = false; throw new Error('withdraw refused'); }
        if (replay) doc.read(probe);
        replay = false;
        return owner.session.withdrawProposals(request);
      },
    } });
    try {
      const alpha = doc.grep({ query: 'alpha' }).matches[0];
      const beta = doc.grep({ query: 'beta' }).matches[0];
      probe = alpha.ref;
      const proposal = await doc.proposeAsync({ author: 'Agent', edits: [{ match: alpha.match, newText: 'ALPHA' }] });
      if (scenario === 'rollback') {
        await expect(doc.accept(proposal.id)).rejects.toThrow('withdraw refused');
        expect(owner.session.getProposals().proposals[0].state).toBe('proposed');
        await doc.accept(proposal.id);
      } else await doc.accept(proposal.id);
      expect(doc.review(proposal.id).status).toBe('accepted');
      const next = await doc.proposeAsync({ author: 'Agent', edits: [{ match: beta.match, newText: 'BETA' }] });
      await doc.rejectAsync(next.id);
      expect(doc.read(beta.ref)).toMatchObject({ revision: beta.revision, text: 'ALPHA beta gamma gamma' });
    } finally { doc.close(); owner.close(); }
  });
}

for (const sameGroup of [false, true]) {
  test(`attached DOCX sync decision detects reentrant editor changes, sameGroup=${sameGroup}`, async () => {
    const owner = await proposalRaceFixture();
    const doc = attachDocx(owner.session);
    let armed = false;
    const unsubscribe = owner.session.onProposalChange(snapshot => {
      if (!armed || !snapshot.proposals.some(record => record.state === 'accepted')) return;
      armed = false;
      expect(owner.session.setProposalStates({ expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
        changes: [{ id: snapshot.proposals[0].id, state: 'rejected' }],
      }).ok).toBe(true);
    });
    try {
      const [alpha, beta, gamma] = ['alpha', 'beta', 'gamma'].map(query => doc.grep({ query }).matches[0]);
      const cursor = doc.grep({ query: 'gamma', limit: 1 }).nextCursor!;
      const first = doc.propose({ author: 'Agent', edits: [{ match: alpha.match, newText: 'ALPHA' }, ...(sameGroup ? [{ match: beta.match, newText: 'BETA' }] : [])] });
      const second = sameGroup ? first : doc.propose({ author: 'Agent', edits: [{ match: beta.match, newText: 'BETA' }] });
      armed = true;
      await expect(doc.accept(second.id)).rejects.toMatchObject({ code: sameGroup ? 'PROPOSAL_DECIDED_IN_EDITOR' : 'STALE_TARGET' });
      expect(owner.session.getProposals().proposals.map(record => record.state)).toEqual(['rejected', 'proposed']);
      expect(() => doc.propose({ author: 'Agent', edits: [{ match: gamma.match, newText: 'GAMMA' }] })).toThrow(DocumentToolError);
      expect(doc.read(gamma.ref).revision).toBeGreaterThan(gamma.revision);
      expect(() => doc.grep({ query: 'gamma', limit: 1, cursor })).toThrow('edited');
    } finally { unsubscribe(); doc.close(); owner.close(); }
  });
}

for (const peer of [false, true]) {
  test(`attached DOCX ignores a lagging snapshot missing a newly registered group, peer=${peer}`, async () => {
    const owner = await proposalRaceFixture();
    let notify: Parameters<typeof owner.session.onProposalChange>[0] | undefined;
    const session = new Proxy(owner.session, { get(target, key) {
      if (key === 'onProposalChange') return (listener: Parameters<typeof owner.session.onProposalChange>[0]) => {
        notify = listener;
        return target.onProposalChange(listener);
      };
      return Reflect.get(target, key);
    } });
    const doc = attachDocx(session, peer ? { editPeer: {
      async proposeChanges(request) { return owner.session.proposeChanges(request); },
      async setProposalStates(request) { return owner.session.setProposalStates(request); },
      async withdrawProposals(request) { return owner.session.withdrawProposals(request); },
    } } : {});
    try {
      const [alpha, beta, gamma] = ['alpha', 'beta', 'gamma'].map(query => doc.grep({ query }).matches[0]);
      const first = await doc.proposeAsync({ author: 'Agent', edits: [{ match: alpha.match, newText: 'ALPHA' }] });
      const older = owner.session.getProposals();
      const second = await doc.proposeAsync({ author: 'Agent', edits: [{ match: beta.match, newText: 'BETA' }] });
      expect(owner.session.getProposals().previewVersion).toBe(older.previewVersion);
      if (!notify) throw new Error('Proposal listener was not registered.');
      notify(older);
      expect(doc.review(second.id)).toMatchObject({ status: 'pending', records: [{ state: 'proposed' }] });
      await doc.accept(second.id);
      expect(doc.read(gamma.ref).revision).toBe(gamma.revision);
      const next = await doc.proposeAsync({ author: 'Agent', edits: [{ match: gamma.match, newText: 'GAMMA' }] });
      await doc.rejectAsync(next.id);
      await doc.rejectAsync(first.id);
      expect(owner.session.getProposals().proposals).toHaveLength(0);
    } finally { doc.close(); owner.close(); }
  });
}
