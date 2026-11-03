import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { attachDocx, attachXlsx, openDocx, openXlsx } from '../src';
import { fixture } from './fixture';
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
  const book = await openXlsx(await xlsxFixture());
  try {
    const cell = book.readCells({ sheet: 'sheet:0', range: 'B3' }).cells[0].cell;
    const proposal = book.proposeCells({ author: 'Agent', edits: [{ cell, input: '123' }] });
    book.handle.editCell(0, 100, 0, 'Unrelated');
    expect(book.review(proposal.id).stale).toBe(false);
    await book.accept(proposal.id);
    expect(book.handle.cell(0, 2, 1).input).toBe('123');
  } finally { book.close(); }
});

test('attached DOCX proposals show engine revisions while pending reads and exports retain original text', async () => {
  const owner = await openDocx(await fixture());
  const agent = attachDocx(owner.session);
  try {
    const [hit] = agent.grep({ query: 'Executive summary' }).matches;
    const before = agent.read(hit.ref);
    const proposal = agent.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] });
    expect(owner.session.getProposals().proposals).toHaveLength(1);
    expect(owner.session.listRevisions().every(revision => revision.author === 'Agent')).toBe(true);
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
    const doc = await openDocx(await fixture());
    try {
      const [hit] = doc.grep({ query: 'Executive summary' }).matches;
      const proposal = doc.propose({ author: 'Agent', edits: [{ match: hit.match, newText: 'Updated summary' }] });
      const revisionIds = doc.session.listRevisions().map(revision => revision.revisionId);
      const dates = doc.session.listRevisions().map(revision => revision.date);
      expect(revisionIds.length).toBeGreaterThan(0);
      expect(await doc.accept(proposal.id, { tracked, ...(tracked ? { date: '2026-01-01T00:00:00Z' } : {}) })).toMatchObject({ id: proposal.id, status: 'accepted', tracked, changes: 1 });
      expect(doc.session.getProposals().proposals).toHaveLength(0);
      expect(doc.session.listRevisions().map(revision => revision.revisionId)).toEqual(tracked ? revisionIds : []);
      if (tracked) expect(doc.session.listRevisions().map(revision => revision.date)).toEqual(dates);
      const zip = await JSZip.loadAsync(await doc.export());
      const xml = await zip.file('word/document.xml')!.async('string');
      expect(xml).toContain('Updated');
      expect(/<w:ins\b/.test(xml)).toBe(tracked);
      expect(doc.review(proposal.id).status).toBe('accepted');
    } finally { doc.close(); }
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
  } finally { agent.close(); edits.dispose(); peer.dispose(); await session.dispose(); }
}, 10000);

test('DOCX proposals on separate ranges preserve original agent offsets', async () => {
  const doc = await openDocx(await fixture());
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
  } finally { doc.close(); }
});

test('DOCX pending render uses a private original view and selected preview', async () => {
  const rendered: string[] = [];
  const doc = await openDocx(await fixture(), { renderer: async session => {
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
  } finally { doc.close(); }
});

test('XLSX proposal notifications refresh a live subscriber without document edits', async () => {
  const book = await openXlsx(await xlsxFixture());
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
  } finally { off(); book.close(); }
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
