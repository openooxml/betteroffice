import { afterAll, beforeAll, expect, setSystemTime, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import {
  EditorDirtyStories,
  dirtyProjectionStory,
  hostSaveMetadata,
  mergeDocxHostMetadata,
  saveEditorDocument,
  serialWorkerSaves,
} from './editorSave';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import type { DocxProposalInput, DocxProposalResult } from './proposals';
import { ResidentEngineWorkerClient, type ResidentProposalReply } from './residentEngineWorkerClient';
import type { ResidentProposalOperation } from './residentEngineWorkerProtocol';
import { yrsToDocument } from './yrsToDocument';

/*
 * Oracle at d6c88fde (0.4.2): DocxEditor/hooks/useYrsCoreSession.ts:1049-1087
 * and DocxEditor/hooks/useFileIO.ts:135-139, under the components/ root below.
 * Dirty call sites at fadfbe49, under packages/docx-react/src/components/:
 * DocxEditor/YrsInput.tsx:613,662,691,514 -> DocxEditor/PagedEditor.tsx:929,1092;
 * DocxEditor.tsx:2083-2086 (reply: host only), 2108-2109 (delete: selection),
 * 2127-2137 (add: selection; viewer: host only at 2121),
 * DocxEditor/editorBatches.ts:208-220 (proposals: new/changed stories),
 * DocxEditor/hooks/usePagedEditorRefApi.ts:209,218 (undo/redo: historyStories).
 * Projection roots/empty-set fallback: DocxEditor/hooks/useYrsCoreSession.ts:1060-1098.
 */

async function saveWorkerArm(arms: Arms, oracle: () => Promise<Uint8Array>) {
  if (arms.peer) return arms.workerSaves!(async (stories) => {
    arms.log.push(`worker save stories=${JSON.stringify([...stories].sort())}`);
    const actual = await arms.client.save({
      comments: hostComments(arms.workerHost),
      host: hostSaveMetadata(arms.workerHost),
      stateVector: arms.peer!.encodeStateVector(), stories,
    });
    const expected = await oracle();
    if (actual.updates.length !== 1) throw new Error('Editor save did not return exactly one diff');
    arms.adoptingWorkerSaveUpdates = true;
    try {
      arms.editorStories.adoptWorkerSaveUpdates(() => {
        for (const update of actual.updates) arms.peer!.applyUpdate(update);
      });
    } finally {
      arms.adoptingWorkerSaveUpdates = false;
    }
    return { actual, expected };
  });
  const peerlessSave = async (_stories: string[]) => {
    arms.log.push('worker save stories=resident');
    if (arms.replica) {
      expect(arms.replica.hasStory('body')).toBe(false);
      expect(arms.worker.requests.some((type) => type === 'encodeState' || type === 'applyUpdate')).toBe(false);
    }
    const actual = await arms.client.save({
      comments: hostComments(arms.workerHost),
      host: hostSaveMetadata(arms.workerHost),
    });
    if (actual.updates.length !== 0) throw new Error('Peerless save returned peer updates');
    return { actual, expected: await oracle() };
  };
  return arms.workerSaves!(peerlessSave);
}

const DATE = '2026-10-02T12:00:00Z';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const STORIES = ['body', 'hf:rIdH1', 'fn:1'] as const;
type Story = (typeof STORIES)[number];
type Topology = 'A/editor' | 'B/viewer' | 'C/hydrate';
type HydrationMode = 'viewer' | 'unloaded' | 'loading';
type Action =
  | 'insert' | 'split' | 'delete' | 'addComment' | 'reply' | 'deleteComment' | 'workerDeleteComment'
  | 'proposal' | 'decide' | 'withdraw' | 'flush' | 'project' | 'undo' | 'redo' | 'save' | 'overlapSave'
  | 'hydrate' | 'hydrateSave' | 'hostChange';
interface Operation {
  action: Action;
  story: Story;
  autoSave?: false;
}

const SEEDS = integerEnv('BO_DIFF_SEEDS', 200);
const OPS = Math.max(20, integerEnv('BO_DIFF_OPS', 32));
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});
let startWorker: (clientId?: number) => InProcessResidentWorker;
let source: Uint8Array;
let font: Uint8Array;

beforeAll(async () => {
  setSystemTime(new Date(DATE));
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'
  ))));
  font = new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
  )));
  startWorker = await residentWorkerFactory();
  source = fixture();
});

afterAll(() => setSystemTime());

function integerEnv(name: string, fallback: number): number {
  const value = Number(process.env[name] ?? fallback);
  if (!Number.isSafeInteger(value) || value < 1) throw new Error(`${name} must be a positive integer`);
  return value;
}

class Random {
  constructor(private state: number) {}

  int(limit: number): number {
    this.state = (this.state + 0x6d2b79f5) >>> 0;
    let value = this.state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return Math.floor(((value ^ (value >>> 14)) >>> 0) / 0x100000000 * limit);
  }

  pick<T>(values: readonly T[]): T {
    if (values.length === 0) throw new Error('Cannot choose from an empty operation target set');
    return values[this.int(values.length)]!;
  }

  shuffle<T>(values: T[]): T[] {
    for (let index = values.length - 1; index > 0; index -= 1) {
      const other = this.int(index + 1);
      [values[index], values[other]] = [values[other]!, values[index]!];
    }
    return values;
  }
}

function run(text: string): string {
  return `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

function paragraph(id: string, content: string): string {
  return `<w:p w14:paraId="${id}">${content}</w:p>`;
}

function fixture(): Uint8Array {
  const parts: Record<string, string> = {
    '[Content_Types].xml':
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="${OFFICE}.header+xml"/>` +
      `<Override PartName="/word/footnotes.xml" ContentType="${OFFICE}.footnotes+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/></Types>`,
    '_rels/.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdDoc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/>` +
      `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/></Relationships>`,
    'word/document.xml':
      `<w:document ${NS} xmlns:r="${REL}"><w:body>` +
      paragraph('0000B001', '<w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/>' +
        '<w:r><w:commentReference w:id="1"/><w:t>Alpha beta gamma.</w:t></w:r>') +
      paragraph('0000B002', run('Delta epsilon zeta.')) +
      paragraph('0000B003', run('Raw prefix ') +
        '<x:inline xmlns:x="urn:bo-diff"><w:r><x:mark x:keep="yes"/></w:r></x:inline>' +
        run(' raw tail.')) +
      paragraph('0000B004', run('Note reference') + '<w:r><w:footnoteReference w:id="1"/></w:r>') +
      '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr></w:body></w:document>',
    'word/header1.xml': `<w:hdr ${NS}>` + paragraph('0000A001',
      '<w:commentRangeStart w:id="2"/>' + run('  Header text  ') +
      '<w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r>') + '</w:hdr>',
    'word/footnotes.xml': `<w:footnotes ${NS}><w:footnote w:id="1">` +
      paragraph('0000F001', run('Footnote text.')) + '</w:footnote></w:footnotes>',
    'word/comments.xml': `<w:comments ${NS}>` +
      `<w:comment w:id="1" w:author="Source" w:date="${DATE}">` +
      paragraph('0000C001', run('Zero-length body comment')) + '</w:comment>' +
      `<w:comment w:id="2" w:author="Source" w:date="${DATE}">` +
      paragraph('0000C002', run('Header comment')) + '</w:comment></w:comments>',
  };
  return new Uint8Array(rezipPartsToArrayBuffer(
    new Map(Object.entries(parts).map(([name, xml]) => [name, toBytes(xml)]))
  ));
}

function hostComments(host: Document): Comment[] {
  return host.package.document.comments ?? [];
}

class MainArm {
  readonly dirtyStories = new Set<string>();
  private compatibilityBase: Document | null = null;

  constructor(
    readonly session: YrsSession,
    readonly host: Document,
    private readonly markWorkerStory?: (story: string) => void
  ) {}

  publishDirectInput(stories?: readonly string[]): void {
    if (!this.session.hasStory('body')) return;
    for (const story of stories ?? [this.session.selection()?.head.story ?? 'body']) {
      const root = ['hf:', 'fn:', 'en:'].some((prefix) => story.startsWith(prefix))
        ? story.split(':', 2).join(':')
        : 'body';
      this.dirtyStories.add(root);
      this.markWorkerStory?.(root);
    }
  }

  project(): Document {
    const base = this.compatibilityBase ?? this.session.materializeDocx();
    if (!base) throw new Error('The main arm has no compatibility package');
    const projected = yrsToDocument(
      this.session, mergeDocxHostMetadata(base, this.host),
      this.dirtyStories.size > 0 ? { storyIds: new Set(this.dirtyStories) } : undefined
    );
    this.dirtyStories.clear();
    this.compatibilityBase = projected;
    return projected;
  }

  async save(): Promise<Uint8Array> {
    const projected = this.project();
    const bytes = await saveEditorDocument(this.session, projected, hostComments(this.host));
    projected.originalBuffer = bytes;
    return new Uint8Array(bytes);
  }

  proposal(call: () => DocxProposalResult): DocxProposalResult {
    const known = new Set(this.session.getProposals().proposals.map(({ id }) => id));
    const since = this.session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
    const result = call();
    if (!result.ok) return result;
    const stories = new Set([
      ...result.snapshot.proposals.filter((proposal) => proposal.changed && !known.has(proposal.id))
        .map((proposal) => proposal.paragraph.story),
      ...this.session.storiesChangedSince(since).stories,
    ]);
    if (stories.size > 0) this.publishDirectInput([...stories]);
    return result;
  }
}

interface Arms {
  worker: InProcessResidentWorker;
  client: ResidentEngineWorkerClient;
  workerHost: Document;
  main: MainArm;
  editorStories: EditorDirtyStories;
  workerSaves?: ReturnType<typeof serialWorkerSaves>;
  peer?: YrsSession;
  replica?: YrsSession;
  hydration?: Promise<void>;
  unsubscribe?: () => void;
  adoptingWorkerSaveUpdates: boolean;
  handbackRemoteUpdates: number;
  otherRemoteUpdates: number;
  mirror?: ResidentProposalReply['mirror'];
  commentStories: Map<number, Story>;
  nextComment: number;
  nextProposal: number;
  log: string[];
}

async function openArms(seed: number, topology: Topology, log: string[]): Promise<Arms> {
  const clientId = 90000 + seed * 2;
  const generation = `resident-save-diff-${seed}`;
  const worker = startWorker(clientId);
  const client = new ResidentEngineWorkerClient(worker);
  let session: YrsSession | undefined;
  let replica: YrsSession | undefined;
  try {
    const { hostJson } = await client.open(source, { generation });
    session = await createYrsSession({ clientId: clientId + (topology === 'A/editor' ? 1 : 0) });
    const { document: host } = session.openDocx(source.slice(), topology !== 'A/editor', { generation });
    if (topology === 'A/editor') {
      session.loadState(await client.encodeState());
      session.beginUndoCapture();
    }
    if (topology === 'C/hydrate') replica = await createYrsSession({ clientId: clientId + 1 });
    session.setSelection({ story: 'body', paraId: '0000B002', offset: 0 });
    if (session.listComments().some(({ id }) => id === '1')) {
      throw new Error('The zero-length source comment unexpectedly has seeded anchors');
    }
    if (!session.resolveComment('2').some(({ story }) => story === 'hf:rIdH1')) {
      throw new Error('The source header comment has no seeded header range');
    }
    const materialized = JSON.stringify(session.materializeDocx());
    if (!materialized.includes('"type":"rawXml"') || !materialized.includes('x:mark')) {
      throw new Error('The foreign inline run was not retained as raw XML');
    }
    const editorStories = new EditorDirtyStories();
    const arms: Arms = {
      worker, client, log,
      workerHost: decodeDocxHostJson(hostJson, source).document,
      main: new MainArm(session, host, topology === 'A/editor'
        ? (story) => editorStories.add(story) : undefined),
      editorStories,
      replica,
      workerSaves: serialWorkerSaves(editorStories),
      adoptingWorkerSaveUpdates: false, handbackRemoteUpdates: 0, otherRemoteUpdates: 0,
      ...(topology === 'A/editor' ? { peer: session } : {}),
      commentStories: new Map<number, Story>([[1, 'body'], [2, 'hf:rIdH1']]),
      nextComment: 3, nextProposal: 1,
    };
    if (arms.peer) listenToPeer(arms);
    return arms;
  } catch (error) {
    client.destroy();
    replica?.destroy();
    session?.destroy();
    throw error;
  }
}

function listenToPeer(arms: Arms): void {
  const session = arms.peer!;
  // PagedEditor remote listener -> publishDirectInput(undefined).
  arms.unsubscribe = session.onUpdate((_update, origin) => {
    if (origin !== 'remote') return;
    if (arms.adoptingWorkerSaveUpdates) arms.handbackRemoteUpdates += 1;
    else arms.otherRemoteUpdates += 1;
    if (!session.hasStory('body')) return;
    arms.editorStories.add(dirtyProjectionStory(session.selection()?.head.story ?? 'body'));
  });
}

function hydrateArms(arms: Arms): Promise<void> {
  if (arms.hydration) return arms.hydration;
  const replica = arms.replica;
  if (!replica || arms.peer) throw new Error('Hydration requires an unloaded editing copy');
  arms.log.push('start editing copy hydration');
  const state = arms.mirror
    ? arms.client.handOver().then((handover) => ({ state: handover.state, mirror: handover }))
    : arms.client.encodeState().then((state) => ({ state, mirror: null }));
  arms.hydration = state.then((handover) => {
    replica.openDocx(source.slice(), false);
    replica.loadState(handover.state);
    if (handover.mirror) {
      replica.mirrorWorkerDocument({ version: handover.mirror.version, proposals: handover.mirror.proposals });
      replica.mirrorWorkerDocument(null);
    }
    for (const story of STORIES) {
      expect(replica.paragraphs(story)).toEqual(arms.main.session.paragraphs(story));
      expect(replica.storyChecksum(story)).toBe(arms.main.session.storyChecksum(story));
      for (const view of ['accepted', 'original'] as const) {
        const actual = replica.readParagraphs({ story, view });
        const expected = arms.main.session.readParagraphs({ story, view });
        if (!actual.ok || !expected.ok) throw new Error('Cannot read the hydrated story');
        expect(actual.paragraphs).toEqual(expected.paragraphs);
      }
    }
    expect(replica.listComments()).toEqual(arms.main.session.listComments());
    for (const { id } of replica.listComments()) {
      expect(replica.resolveComment(id)).toEqual(arms.main.session.resolveComment(id));
    }
    expect(replica.getProposals().proposals).toEqual(arms.main.session.getProposals().proposals);
    expect(replica.getProposals().previewVersion).toBe(arms.main.session.getProposals().previewVersion);
    expect(hostComments(arms.workerHost)).toEqual(hostComments(arms.main.host));
    const selection = arms.main.session.selection();
    if (selection) replica.setSelection(selection.anchor, selection.head);
    arms.peer = replica;
    listenToPeer(arms);
    replica.beginUndoCapture();
    arms.log.push(`editing copy hydrated clientId=${replica.clientId}`);
  });
  return arms.hydration;
}

function editPeer<T>(arms: Arms, edit: (session: YrsSession) => T): T {
  const peer = arms.peer!;
  const updates: Uint8Array[] = [];
  const unsubscribe = peer !== arms.main.session
    ? peer.onUpdate((update, origin) => {
        if (origin === 'local') updates.push(update);
      })
    : undefined;
  try {
    const result = edit(peer);
    for (const update of updates) arms.main.session.applyLocalUpdate(update);
    return result;
  } finally {
    unsubscribe?.();
  }
}

function publishPeerInput(arms: Arms, stories?: readonly string[]): void {
  if (arms.peer && arms.peer !== arms.main.session) {
    const selection = arms.peer.selection();
    if (selection) arms.main.session.setSelection(selection.anchor, selection.head);
    for (const story of stories ?? [selection?.head.story ?? 'body']) {
      arms.editorStories.add(dirtyProjectionStory(story));
    }
  }
  arms.main.publishDirectInput(stories);
}

async function bootstrap(arms: Arms): Promise<void> {
  await arms.client.bootstrap({
    workerAuthoritative: true,
    clientId: 0,
    state: new Uint8Array(0),
    selection: null,
    fonts: [font], fontsRevision: 1,
    renderInputs: [], measureInputs: [],
    layoutInput: LAYOUT, layoutWithRegions: true, layoutRevision: 1,
  }, '{}', { opened: true, layoutExtras: '{}' });
}

async function flushPeer(arms: Arms): Promise<void> {
  if (!arms.peer) return;
  const update = arms.peer.encodeStateAsUpdate(arms.client.remoteStateVector() ?? undefined);
  arms.client.invalidate(update, arms.peer.selection());
  await arms.client.revisionCount();
  arms.log.push(`flush peer diff (${update.byteLength} bytes)`);
}

async function workerMutation(
  arms: Arms, operation: ResidentProposalOperation
): Promise<ResidentProposalReply> {
  const known = new Set(arms.mirror?.proposals.entries.map(({ record }) => record.id));
  const reply = await arms.client.proposal(operation);
  if (operation.kind !== 'snapshot' && operation.kind !== 'removeComment') {
    if (!reply.result) throw new Error('Worker proposal reply omitted its result');
    if (!reply.result.ok) arms.log.push(`worker proposal refused ${JSON.stringify(reply.result)}`);
  }
  if (arms.peer) {
    for (const update of reply.updates) arms.peer.applyHostUpdate(update, reply.changedStories);
    if (operation.kind === 'removeComment') {
      if (reply.changedStories.length > 0) arms.main.publishDirectInput();
    } else {
      const stories = new Set([
        ...reply.mirror.proposals.entries.filter(({ record }) => record.changed && !known.has(record.id))
          .map(({ record }) => record.paragraph.story),
        ...reply.changedStories,
      ]);
      if (stories.size > 0) arms.main.publishDirectInput([...stories]);
    }
  }
  if (!arms.peer) {
    for (const story of reply.projectionStories ?? []) arms.editorStories.add(story);
  }
  arms.mirror = reply.mirror;
  return reply;
}

function target(arms: Arms, story: Story, random: Random) {
  const paragraphs = (arms.peer ?? arms.main.session).paragraphs(story);
  const editable = story === 'body' ? paragraphs.filter(({ paraId }) => paraId !== '0000B004') : paragraphs;
  const paragraph = random.pick(editable);
  const offset = random.int(paragraph.text.length + 1);
  const at = { story, paraId: paragraph.paraId, offset };
  return { at, paragraph };
}

function focus(arms: Arms, story: Story, random: Random) {
  const chosen = target(arms, story, random);
  arms.main.session.setSelection(chosen.at);
  if (arms.peer && arms.peer !== arms.main.session) arms.peer.setSelection(chosen.at);
  return chosen;
}

function mainProposal(arms: Arms, call: () => DocxProposalResult): void {
  const result = arms.main.proposal(call);
  if (!result.ok) arms.log.push(`main proposal refused ${JSON.stringify(result)}`);
}

function changeComments(arms: Arms, change: (comments: Comment[]) => Comment[]): void {
  for (const host of [arms.main.host, arms.workerHost]) {
    host.package.document.comments = change(hostComments(host));
  }
}

function commentBody(text: string): Comment['content'] {
  return [{ type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text }] }] }];
}

function addComment(arms: Arms, story: Story, random: Random): number {
  const { at, paragraph } = arms.peer ? focus(arms, story, random) : target(arms, story, random);
  const comment: Comment = {
    id: arms.nextComment++, author: 'Differential', date: DATE,
    content: commentBody(`Comment ${arms.nextComment - 1} in ${story}`),
  };
  if (arms.peer) {
    const span = arms.peer.locateParagraph(story, at.paraId);
    const start = span.start + Math.min(at.offset, Math.max(0, paragraph.text.length - 1));
    editPeer(arms, (session) => session.applyRawOps(story, [{
      op: 'setComment', id: String(comment.id), ranges: [[start, start + 1]],
      author: comment.author, date: comment.date, body: comment.content,
    }]));
    publishPeerInput(arms);
  }
  changeComments(arms, (comments) => [...comments, structuredClone(comment)]);
  arms.commentStories.set(comment.id, story);
  arms.log.push(
    `add comment ${comment.id} ${story}/${at.paraId}:${at.offset} ` +
    (arms.peer ? 'peer anchor' : 'viewer host only')
  );
  return comment.id;
}

function parentComment(arms: Arms, story: Story, random: Random): number {
  const parents = hostComments(arms.main.host)
    .filter(({ id, parentId }) => parentId == null && arms.commentStories.get(id) === story);
  return parents.length > 0 ? random.pick(parents).id : addComment(arms, story, random);
}

async function textProposal(
  arms: Arms, story: Story, random: Random, text: string, deleteText = false
): Promise<void> {
  const { at, paragraph } = target(arms, story, random);
  const accepted = arms.main.session.readParagraphs({
    story, paraIds: [paragraph.paraId], view: 'accepted',
  });
  if (!accepted.ok) throw new Error(`Cannot read proposal target: ${JSON.stringify(accepted)}`);
  const acceptedText = accepted.paragraphs[0]?.text;
  if (acceptedText === undefined) throw new Error('The proposal target paragraph is missing');
  const id = `diff-${arms.nextProposal++}`;
  const input: DocxProposalInput = {
    id,
    paragraph: {
      kind: 'persisted',
      story: story === 'body' ? { kind: 'body', partUri: '/word/document.xml' }
        : story === 'hf:rIdH1' ? { kind: 'header', partUri: '/word/header1.xml' }
        : { kind: 'footnote', partUri: '/word/footnotes.xml', itemId: '1' },
      paraId: paragraph.paraId,
    },
    suggest: { author: `Differential ${id}`, date: DATE },
    ...(deleteText && acceptedText.length > 0
      ? {
          op: 'replaceText' as const, search: acceptedText.slice(0, 1),
          replaceWith: '', occurrence: 'first' as const,
        }
      : { op: 'insertText' as const, at: 'end' as const, text }),
  };
  arms.log.push(`worker proposal ${JSON.stringify(input)} (selection ${at.story})`);
  await flushPeer(arms);
  const snapshot = await arms.client.proposal({ kind: 'snapshot' });
  arms.mirror = snapshot.mirror;
  await workerMutation(arms, {
    kind: 'propose', request: { expectVersion: snapshot.mirror.version, proposals: [input] },
  });
  if (!arms.peer) mainProposal(arms, () => arms.main.session.proposeChanges({
    expectVersion: arms.main.session.version(), proposals: [input],
  }));
}

async function decideOrWithdraw(arms: Arms, random: Random, withdraw: boolean): Promise<void> {
  await flushPeer(arms);
  const snapshot = await arms.client.proposal({ kind: 'snapshot' });
  arms.mirror = snapshot.mirror;
  const entries = snapshot.mirror.proposals.entries;
  if (entries.length === 0) {
    const stories = arms.replica && !arms.peer ? ['hf:rIdH1', 'fn:1'] as const : STORIES;
    await textProposal(arms, random.pick(stories), random, ' proposed');
    return;
  }
  const id = random.pick(entries).record.id;
  const state = random.pick(['accepted', 'rejected', 'proposed'] as const);
  arms.log.push(`${withdraw ? 'withdraw' : `decide ${state}`} proposal ${id}`);
  await workerMutation(arms, withdraw
    ? { kind: 'withdraw', request: { expectVersion: snapshot.mirror.version, ids: [id] } }
    : { kind: 'setStates', request: {
        expectVersion: snapshot.mirror.version,
        expectPreviewVersion: snapshot.mirror.proposals.previewVersion, changes: [{ id, state }],
      } });
  if (!arms.peer) mainProposal(arms, () => withdraw
    ? arms.main.session.withdrawProposals({ expectVersion: arms.main.session.version(), ids: [id] })
    : arms.main.session.setProposalStates({
        expectVersion: arms.main.session.version(),
        expectPreviewVersion: arms.main.session.getProposals().previewVersion, changes: [{ id, state }],
      }));
}

async function applyOperation(arms: Arms, operation: Operation, random: Random): Promise<void> {
  const { action, story } = operation;
  const session = arms.peer ?? arms.main.session;
  switch (action) {
    case 'split': {
      if (!arms.peer) {
        await applyOperation(arms, { action: 'insert', story }, random);
        return;
      }
      const { at, paragraph } = focus(arms, story, random);
      const offset = random.int(paragraph.text.length + 1);
      arms.log.push(`peer split ${story}/${at.paraId}:${offset}`);
      editPeer(arms, (session) => {
        session.addUndoBoundary();
        const receipt = session.splitParagraph({ ...at, offset });
        if (receipt.secondParaId) session.setSelection({ story, paraId: receipt.secondParaId, offset: 0 });
      });
      publishPeerInput(arms);
      return;
    }
    case 'insert':
    case 'delete': {
      const text = random.pick([' x', 'Y', '  z ', 'q&']);
      if (!arms.peer) {
        await textProposal(arms, story, random, text, action === 'delete');
        return;
      }
      const { at, paragraph } = focus(arms, story, random);
      editPeer(arms, (session) => {
        session.addUndoBoundary();
        if (action === 'delete' && paragraph.text.length > 0) {
          const start = Math.min(at.offset, paragraph.text.length - 1);
          const end = Math.min(paragraph.text.length, start + 1 + random.int(3));
          arms.log.push(`peer delete ${story}/${at.paraId} [${start},${end})`);
          const receipt = session.deleteRange({
            story,
            start: { paraId: at.paraId, offset: start },
            end: { paraId: at.paraId, offset: end },
          });
          session.setSelection(receipt.range ? { story, ...receipt.range.start } : { ...at, offset: start });
        } else {
          arms.log.push(`peer insert ${story}/${at.paraId}:${at.offset} ${JSON.stringify(text)}`);
          const receipt = session.insertText(at, text);
          session.setSelection(
            receipt.range ? { story, ...receipt.range.end } : { ...at, offset: at.offset + text.length }
          );
        }
      });
      publishPeerInput(arms);
      return;
    }
    case 'addComment':
      addComment(arms, story, random);
      return;
    case 'reply': {
      const parentId = parentComment(arms, story, random);
      const reply: Comment = {
        id: arms.nextComment++, parentId, author: 'Reply', date: DATE,
        content: commentBody(`Reply to ${parentId}`),
      };
      arms.log.push(`reply ${reply.id} to ${parentId} ${story} (host only)`);
      changeComments(arms, (comments) => [...comments, structuredClone(reply)]);
      return;
    }
    case 'deleteComment':
    case 'workerDeleteComment': {
      const id = parentComment(arms, story, random);
      const selected = arms.peer
        ? focus(arms, random.pick(STORIES), random).at.story
        : session.selection()?.head.story ?? 'body';
      changeComments(arms, (comments) => comments.filter(
        (comment) => comment.id !== id && comment.parentId !== id
      ));
      const authority = arms.peer && action === 'deleteComment' ? 'peer' : 'worker';
      arms.log.push(`${authority} removeComment ${id} ${story} (selection ${selected})`);
      if (!arms.peer || action === 'workerDeleteComment') {
        await flushPeer(arms);
        await workerMutation(arms, { kind: 'removeComment', id: String(id) });
      }
      if (arms.peer && action === 'deleteComment') {
        try {
          editPeer(arms, (session) => session.applyRawOps('body', [{ op: 'removeComment', id: String(id) }]));
          publishPeerInput(arms);
        } catch {}
      } else if (!arms.peer) {
        try {
          arms.main.session.applyRawOps('body', [{ op: 'removeComment', id: String(id) }]);
          arms.main.publishDirectInput();
        } catch {}
      }
      return;
    }
    case 'proposal':
      await textProposal(arms, story, random, random.pick([' proposed', ' P ', 'Q']));
      return;
    case 'decide':
    case 'withdraw':
      await decideOrWithdraw(arms, random, action === 'withdraw');
      return;
    case 'flush':
      if (arms.peer) await flushPeer(arms);
      else arms.log.push('flush (viewer has no peer)');
      return;
    case 'project':
      arms.log.push(`main project dirty=${JSON.stringify([...arms.main.dirtyStories].sort())}`);
      arms.main.project();
      arms.editorStories.projected();
      return;
    case 'undo':
    case 'redo': {
      const changed = editPeer(arms, (session) => action === 'undo' ? session.undo() : session.redo());
      const stories = changed ? session.historyStories() : [];
      arms.log.push(`${action} ${changed} ${JSON.stringify(stories)}`);
      if (changed) publishPeerInput(arms, stories);
      return;
    }
    case 'hostChange': {
      for (const host of [arms.main.host, arms.workerHost]) {
        const body = host.package.document;
        const properties = body.finalSectionProperties;
        body.finalSectionProperties = {
          ...properties, marginTop: (properties?.marginTop ?? 1440) + 120,
        };
      }
      arms.log.push('host margin change (no story marks)');
      return;
    }
  }
}

function operations(topology: Topology, random: Random, hydrationMode?: HydrationMode): Operation[] {
  const group = (story: Story, actions: Action[]): Operation[] =>
    actions.map((action) => ({ action, story }));
  if (topology === 'C/hydrate') {
    const peerless = random.shuffle([
      group('hf:rIdH1', hydrationMode === 'viewer'
        ? ['insert', 'delete', 'addComment'] : ['insert', 'delete']),
      group('fn:1', ['proposal', 'decide', 'withdraw']),
    ]).flat();
    const transition = [
      ...group('hf:rIdH1', ['proposal', 'save']),
      ...group('body', hydrationMode === 'loading'
        ? ['hydrateSave', 'save'] : ['hydrate', 'hostChange', 'save', 'save']),
    ];
    const workerOnly: Action[] = ['proposal', 'decide', 'withdraw', 'workerDeleteComment'];
    return [
      ...[...peerless, ...transition].map((operation) => ({ ...operation, autoSave: false as const })),
      ...operations('A/editor', random).filter(({ action }) => !workerOnly.includes(action)),
    ];
  }
  const groups = [
    ...STORIES.map((story) => group(story, ['insert', 'delete', 'proposal'])),
    ...(['body', 'hf:rIdH1'] as const).map((story) => group(story, ['addComment', 'reply', 'deleteComment'])),
    group('body', ['workerDeleteComment']), group('hf:rIdH1', ['workerDeleteComment']),
    group('body', ['flush', 'decide', 'withdraw']),
  ];
  if (topology === 'A/editor') groups.push(
    [
      { action: 'insert', story: 'body' },
      { action: 'project', story: 'body' },
      { action: 'insert', story: 'hf:rIdH1' },
    ],
    group(random.pick(STORIES), ['insert', 'undo', 'redo']),
    group(random.pick(STORIES), ['split', 'save', 'save']),
    group(random.pick(STORIES), ['split', 'overlapSave']),
    group(random.pick(STORIES), ['save', 'save'])
  );
  const required = random.shuffle(groups).flat();
  const weighted: Action[] = ['insert', 'insert', 'delete', 'delete', 'addComment', 'addComment',
    'reply', 'reply', 'deleteComment', 'workerDeleteComment', 'proposal', 'decide', 'withdraw', 'flush', 'save',
    ...(topology === 'A/editor' ? ['project', 'split', 'overlapSave', 'undo', 'redo'] as const : [])];
  while (required.length < OPS) {
    const action = random.pick(weighted);
    required.push({
      action,
      story: random.pick(action.includes('Comment') || action === 'reply'
        ? ['body', 'hf:rIdH1'] as const : STORIES),
    });
  }
  return required;
}

function partDifferences(worker: Uint8Array, main: Uint8Array): Array<{ part: string; text: string }> {
  const actual = unzipContainer(worker);
  const expected = unzipContainer(main);
  const differences: Array<{ part: string; text: string }> = [];
  const snippet = (bytes: Uint8Array | undefined, offset: number) => bytes
    ? JSON.stringify(new TextDecoder().decode(bytes.subarray(Math.max(0, offset - 80), offset + 81)))
    : '<missing part>';
  for (const part of [...new Set([...Object.keys(actual), ...Object.keys(expected)])].sort()) {
    const a = actual[part];
    const b = expected[part];
    let offset = 0;
    if (a && b) {
      while (offset < Math.min(a.length, b.length) && a[offset] === b[offset]) offset += 1;
      if (offset === a.length && offset === b.length) continue;
    }
    differences.push({
      part,
      text: `${part} at byte ${offset} (worker=${a?.length ?? 'missing'}, main=${b?.length ?? 'missing'})\n` +
        `  worker ${snippet(a, offset)}\n  main   ${snippet(b, offset)}`,
    });
  }
  return differences;
}

function liveCommentIds(parts: Record<string, Uint8Array>): ReadonlySet<string> {
  const xml = parts['word/comments.xml'] ? new TextDecoder().decode(parts['word/comments.xml']) : '';
  return new Set(
    [...xml.matchAll(/<w:comment\b[^>]*\bw:id="([^"]+)"/g)].map((match) => match[1]!)
  );
}

function liveSaveCommentIds(commentsXmlIds: ReadonlySet<string>, saveCommentIds: readonly string[]): ReadonlySet<string> {
  return new Set(saveCommentIds.filter((id) => commentsXmlIds.has(id)));
}

const XML_OPAQUE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/g;

function danglingRangeMarkers(xml: string, liveIds: ReadonlySet<string>): string[] {
  return [...xml.replace(XML_OPAQUE, '').matchAll(/<w:commentRange(?:Start|End)\b[^>]*\bw:id="([^"]+)"/g)]
    .map((match) => match[1]!)
    .filter((id) => !liveIds.has(id));
}

const COMMENT_RANGE_MARKER = /<w:commentRange(?:Start|End)\b[^<>]*\bw:id="([^"]+)"[^<>]*\/>/g;
const COMMENT_REFERENCE_RUN =
  /<w:r\b[^<>]*>\s*(?:<w:rPr\b[^<>]*>(?:(?!<\/?w:r\b|<\/w:rPr>)[\s\S])*<\/w:rPr>\s*)?<w:commentReference\b[^<>]*\bw:id="([^"]+)"[^<>]*\/>\s*<\/w:r>/g;

function withoutStaleCommentMarkers(xml: string, ids: ReadonlySet<string>): string {
  const strip = (text: string) => [COMMENT_REFERENCE_RUN, COMMENT_RANGE_MARKER].reduce((result, pattern) =>
    result.replace(pattern, (marker, id: string) => ids.has(id) ? '' : marker), text
  );
  let result = '';
  let offset = 0;
  for (const opaque of xml.matchAll(XML_OPAQUE)) {
    result += strip(xml.slice(offset, opaque.index)) + opaque[0];
    offset = opaque.index! + opaque[0].length;
  }
  return result + strip(xml.slice(offset));
}

function mergedTextRuns(xml: string): string {
  const pattern = /^<w:r>(<w:rPr>(?:(?!<\/?w:r\b|<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:t( xml:space="preserve")?>([^<]*)<\/w:t><\/w:r>$/;
  let result = '';
  let offset = 0;
  let runStart = 0;
  let runDepth = 0;
  let previous: { start: number; end: number; properties: string; text: string; count: number } | undefined;
  const flush = () => {
    if (!previous) return;
    result += xml.slice(offset, previous.start);
    const space = /^\s|\s$/.test(previous.text) ? ' xml:space="preserve"' : '';
    result += previous.count === 1 ? xml.slice(previous.start, previous.end)
      : `<w:r>${previous.properties}<w:t${space}>${previous.text}</w:t></w:r>`;
    offset = previous.end;
  };
  for (const tag of xml.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<\/?w:r\b[^<>]*>/g)) {
    if (/^<w:r\b/.test(tag[0])) {
      if (runDepth === 0) runStart = tag.index!;
      if (!tag[0].endsWith('/>')) runDepth += 1;
      continue;
    }
    if (tag[0] !== '</w:r>' || runDepth === 0) continue;
    runDepth -= 1;
    if (runDepth !== 0) continue;
    const start = runStart;
    const end = tag.index! + tag[0].length;
    const match = xml.slice(start, end).match(pattern);
    if (!match || match[3]!.includes('&#') || (!match[2] && /^\s|\s$/.test(match[3]!))) continue;
    const properties = match[1] ?? '';
    if (previous && previous.end === start && previous.properties === properties) {
      previous.end = end;
      previous.text += match[3]!;
      previous.count += 1;
    } else {
      flush();
      previous = { start, end, properties, text: match[3]!, count: 1 };
    }
  }
  flush();
  return result + xml.slice(offset);
}

function commentReferenceIds(xml: string): string[] {
  return [...xml.replace(XML_OPAQUE, '').matchAll(/<w:commentReference\b[^>]*\bw:id="([^"]+)"/g)].map((match) => match[1]!);
}

function staleMainPartExempt(workerXml: string, mainXml: string, liveIds: ReadonlySet<string>): boolean {
  const ids = new Set(danglingRangeMarkers(mainXml, liveIds));
  return ids.size > 0 && danglingRangeMarkers(workerXml, liveIds).length === 0 &&
    !commentReferenceIds(workerXml).some((id) => ids.has(id)) &&
    mergedTextRuns(withoutStaleCommentMarkers(mainXml, ids)) === mergedTextRuns(workerXml);
}

// Deferred: main-thread saves can retain deleted comment range markers.
function staleMainCommentMarkers(
  worker: Uint8Array, main: Uint8Array, parts: readonly string[], saveCommentIds: readonly string[], log: string[]
): boolean {
  const actual = unzipContainer(worker);
  const expected = unzipContainer(main);
  const liveIds = liveSaveCommentIds(liveCommentIds(expected), saveCommentIds);
  const matches = parts.length > 0 && parts.every((part) => {
    if (!/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(part) ||
      actual[part] === undefined || expected[part] === undefined) return false;
    const workerXml = new TextDecoder().decode(actual[part]);
    const mainXml = new TextDecoder().decode(expected[part]);
    return staleMainPartExempt(workerXml, mainXml, liveIds);
  });
  if (matches) {
    for (const part of parts) {
      const xml = new TextDecoder().decode(expected[part]);
      const workerXml = new TextDecoder().decode(actual[part]);
      const ids = new Set(danglingRangeMarkers(xml, liveIds));
      const strippedMain = withoutStaleCommentMarkers(xml, ids);
      const merged = mergedTextRuns(strippedMain) !== strippedMain || mergedTextRuns(workerXml) !== workerXml;
      log.push(`stale main-thread comment markers in ${part}: ids=${JSON.stringify([...ids])} ` +
        `comments=${JSON.stringify(saveCommentIds)} merged runs=${merged}`);
    }
  }
  return matches;
}

test('live save comment ids intersect comments XML with the save list', () => {
  const commentsXmlIds = new Set(['1', '5']);
  expect(liveSaveCommentIds(commentsXmlIds, [])).toEqual(new Set<string>());
  expect(liveSaveCommentIds(commentsXmlIds, ['5'])).toEqual(new Set(['5']));
  expect(liveSaveCommentIds(new Set(['1']), ['1', '9'])).toEqual(new Set(['1']));
});

test('the stale comment marker exemption removes only dangling markers', () => {
  const worker = '<w:p><w:r><w:t>text</w:t></w:r></w:p>';
  const range = '<w:commentRangeStart w:id="7"/><w:commentRangeEnd w:id="7"/>';
  const reference = '<w:r><w:commentReference w:id="7"/></w:r>';
  const formattedReference = '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' +
    '<w:commentReference w:id="7"/></w:r>';
  const main = worker.replace('</w:p>', `${range}${reference}${formattedReference}</w:p>`);
  const liveIds = new Set<string>();
  const ids = new Set(['7']);
  expect(staleMainPartExempt(worker, main, liveIds)).toBe(true);
  expect(staleMainPartExempt(worker, main.replace('text', 'text!'), liveIds)).toBe(false);
  expect(staleMainPartExempt(worker, main.replace('</w:p>', '<x:foreign/></w:p>'), liveIds)).toBe(false);
  expect(staleMainPartExempt(worker, main, new Set(['7']))).toBe(false);
  expect(staleMainPartExempt(worker, worker, liveIds)).toBe(false);
  expect(staleMainPartExempt(main, main, liveIds)).toBe(false);
  expect(withoutStaleCommentMarkers(main, ids)).toBe(worker);
  expect(withoutStaleCommentMarkers(main, new Set<string>())).toBe(main);
  const mixed = '<w:r><w:commentReference w:id="7"/><w:t>keep</w:t></w:r>';
  expect(withoutStaleCommentMarkers(mixed, ids)).toBe(mixed);
  expect(staleMainPartExempt(worker, main.replace(reference, mixed), liveIds)).toBe(false);
  const foreign = reference.replace('</w:r>', '<x:foreign/></w:r>');
  expect(withoutStaleCommentMarkers(foreign, ids)).toBe(foreign);
  expect(staleMainPartExempt(worker, main.replace(reference, foreign), liveIds)).toBe(false);
  const mixedMarkers = `<w:r>${range}<w:commentReference w:id="7"/></w:r>`;
  expect(withoutStaleCommentMarkers(mixedMarkers, ids)).toBe(reference);
  const bare = '<w:commentReference w:id="7"/>';
  expect(withoutStaleCommentMarkers(bare, ids)).toBe(bare);
  expect(staleMainPartExempt(worker, main.replace(reference, bare), liveIds)).toBe(false);
  const liveRange = range.replaceAll('"7"', '"8"');
  expect(withoutStaleCommentMarkers(`${main}${liveRange}`, ids)).toBe(`${worker}${liveRange}`);
  expect(staleMainPartExempt(worker, `${main}${liveRange}`, new Set(['8']))).toBe(false);
  expect(staleMainPartExempt(`${worker}${liveRange}`, `${main}${liveRange}`, new Set(['8']))).toBe(true);
});

test('the stale comment marker exemption merges cached text runs without hiding differences', () => {
  const reference = '<w:r><w:commentReference w:id="1"/></w:r>';
  const worker = `<w:p><w:r><w:t>Alpa betaY gamma.</w:t></w:r>${reference}</w:p>`;
  const main = '<w:p><w:r><w:t>Alpa beta</w:t></w:r><w:commentRangeStart w:id="5"/>' +
    '<w:r><w:t>Y</w:t></w:r><w:commentRangeEnd w:id="5"/>' +
    '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="5"/></w:r>' +
    `<w:r><w:t xml:space="preserve"> gamma.</w:t></w:r>${reference}</w:p>`;
  const commentsXmlIds = liveCommentIds({
    'word/comments.xml': toBytes('<w:comments><w:comment w:id="1"/><w:comment w:id="5"/></w:comments>'),
  });
  const liveIds = liveSaveCommentIds(commentsXmlIds, []);
  expect(staleMainPartExempt(worker, main, liveIds)).toBe(true);
  const splitWorker = worker.replace('<w:r><w:t>Alpa betaY gamma.</w:t></w:r>',
    '<w:r><w:t xml:space="preserve">Alpa </w:t></w:r><w:r><w:t>betaY gamma.</w:t></w:r>');
  expect(staleMainPartExempt(splitWorker, main, liveIds)).toBe(true);
  const unpreservedWorker = splitWorker.replace(' xml:space="preserve"', '');
  expect(staleMainPartExempt(unpreservedWorker, main, liveIds)).toBe(false);
  expect(mergedTextRuns(unpreservedWorker)).toBe(unpreservedWorker);
  const encodedMain = main.replace('<w:t>Alpa beta</w:t>', '<w:t xml:space="preserve">&#32;Alpa beta</w:t>');
  const encodedWorker = worker.replace('<w:t>Alpa betaY gamma.</w:t>', '<w:t>&#32;Alpa betaY gamma.</w:t>');
  expect(staleMainPartExempt(encodedWorker, encodedMain, liveIds)).toBe(false);
  const cdata = '<w:r><w:t><![CDATA[<w:commentRangeStart w:id="5"/>]]></w:t></w:r>';
  expect(withoutStaleCommentMarkers(cdata, new Set(['5']))).toBe(cdata);
  const cdataMain = main.replace('</w:p>', `${cdata}</w:p>`);
  expect(staleMainPartExempt(worker.replace('</w:p>', `${cdata}</w:p>`), cdataMain, liveIds)).toBe(true);
  expect(staleMainPartExempt(worker.replace('</w:p>', '<w:r><w:t><![CDATA[]]></w:t></w:r></w:p>'), cdataMain, liveIds))
    .toBe(false);
  for (const opaque of ['<!--keep-->', '<?keep?>']) {
    const opaqueReference = `<w:r><w:rPr>${opaque}</w:rPr><w:commentReference w:id="5"/></w:r>`;
    expect(staleMainPartExempt(worker.replace('</w:p>', `${opaqueReference}</w:p>`),
      main.replace('</w:p>', `${opaqueReference}</w:p>`), liveIds)).toBe(false);
  }
  expect(staleMainPartExempt(worker, main.replace('gamma', 'gamma!'), liveIds)).toBe(false);
  const formatted = main.replace('<w:r><w:t>Y</w:t></w:r>',
    '<w:r><w:rPr><w:b/></w:rPr><w:t>Y</w:t></w:r>');
  expect(staleMainPartExempt(worker, formatted, liveIds)).toBe(false);
  const extraReference = reference.replace('"1"', '"2"');
  expect(staleMainPartExempt(worker, main.replace('</w:p>', `${extraReference}</w:p>`), liveIds)).toBe(false);
  expect(staleMainPartExempt(worker, main.replace('</w:p>', '<w:r><w:tab/></w:r></w:p>'), liveIds)).toBe(false);
  expect(staleMainPartExempt(worker.replace(reference, ''), main, liveIds)).toBe(false);
  expect(staleMainPartExempt(worker, main, liveSaveCommentIds(commentsXmlIds, ['5']))).toBe(false);
});

test('text run merging preserves formatting, whitespace and other children', () => {
  const first = '<w:r><w:t>one</w:t></w:r>';
  const second = '<w:r><w:t>two</w:t></w:r>';
  expect(mergedTextRuns(`${first}${second}${first}`)).toBe('<w:r><w:t>onetwoone</w:t></w:r>');
  const spaced = '<w:r><w:t xml:space="preserve"> two </w:t></w:r>';
  expect(mergedTextRuns(`${spaced}${first}`)).toBe('<w:r><w:t xml:space="preserve"> two one</w:t></w:r>');
  expect(mergedTextRuns(`${first}${spaced}`)).toBe('<w:r><w:t xml:space="preserve">one two </w:t></w:r>');
  expect(mergedTextRuns(`${first}${spaced}${second}`)).toBe('<w:r><w:t>one two two</w:t></w:r>');
  expect(mergedTextRuns(spaced)).toBe(spaced);
  const properties = '<w:rPr><w:b/></w:rPr>';
  const boldFirst = first.replace('<w:r>', `<w:r>${properties}`);
  const boldSecond = second.replace('<w:r>', `<w:r>${properties}`);
  expect(mergedTextRuns(`${boldFirst}${boldSecond}`)).toBe(`<w:r>${properties}<w:t>onetwo</w:t></w:r>`);
  expect(mergedTextRuns(`${boldFirst}${second}`)).toBe(`${boldFirst}${second}`);
  const otherProperties = boldSecond.replace('<w:b/>', '<w:b />');
  expect(mergedTextRuns(`${boldFirst}${otherProperties}`)).toBe(`${boldFirst}${otherProperties}`);
  for (const separator of [' ', '\n', '<w:bookmarkStart w:id="1"/>', '<!--keep-->', '</w:p><w:p>']) {
    const xml = `<w:p>${first}${separator}${second}</w:p>`;
    expect(mergedTextRuns(xml)).toBe(xml);
  }
  for (const run of [
    '<w:r><w:t> keep</w:t></w:r>',
    '<w:r><w:t>keep </w:t></w:r>',
    '<w:r><w:t> keep </w:t></w:r>',
    '<w:r><w:t xml:space="preserve">&#32;keep</w:t></w:r>',
    '<w:r><w:tab/></w:r>',
    '<w:r><w:t>keep</w:t><w:tab/></w:r>',
    '<w:r><w:t>keep</w:t><w:br/></w:r>',
    '<w:r><w:fldChar w:fldCharType="begin"/><w:t>keep</w:t></w:r>',
    '<w:r><w:instrText>keep</w:instrText></w:r>',
    '<w:r><x:foreign/><w:t>keep</w:t></w:r>',
    '<w:r w:rsidR="1"><w:t>keep</w:t></w:r>',
    '<w:r><w:t xml:lang="en">keep</w:t></w:r>',
    '<w:r><w:t>keep</w:t><w:t>more</w:t></w:r>',
    `<w:r><x:foreign>${first}${second}${first}</x:foreign></w:r>`,
    `<!--${first}${second}-->`,
    `<![CDATA[${first}${second}]]>`,
  ]) {
    const xml = `${first}${run}${second}`;
    expect(mergedTextRuns(xml)).toBe(xml);
  }
});

test('seeded resident DOCX saves match the 0.4.2 main-thread save', async () => {
  const mismatchingSeeds = new Set<number>();
  const staleMainSeeds = new Set<number>();
  const failures: string[] = [];
  let staleMainSaves = 0;
  let erroredSeeds = 0;
  let seedsRun = 0;
  let savesCompared = 0;
  let handbackRemoteUpdates = 0;
  let otherRemoteUpdates = 0;
  let editorSeeds = 0;
  const hydrationSeeds: Record<HydrationMode, number> = { viewer: 0, unloaded: 0, loading: 0 };
  for (let seed = 1; seed <= SEEDS; seed += 1) {
    const random = new Random(seed);
    const topology: Topology = seed % 3 === 0 ? 'C/hydrate'
      : seed % 3 === 1 ? 'A/editor' : 'B/viewer';
    const hydrationMode = topology === 'C/hydrate'
      ? (['viewer', 'unloaded', 'loading'] as const)[(seed / 3 - 1) % 3]!
      : undefined;
    const log: string[] = [`open seed=${seed} topology=${topology}`];
    let arms: Arms | undefined;
    seedsRun += 1;
    if (topology === 'A/editor') editorSeeds += 1;
    if (hydrationMode) {
      hydrationSeeds[hydrationMode] += 1;
      log.push(`peerless ${hydrationMode === 'viewer' ? 'viewer session' : 'editable session, copy unloaded'}`);
    }
    try {
      arms = await openArms(seed, topology, log);
      const current = arms;
      let saveNumber = 0;
      const compareSaved = (
        { actual, expected }: Awaited<ReturnType<typeof saveWorkerArm>>, saveCommentIds: readonly string[]
      ) => {
        savesCompared += 1;
        const differences = partDifferences(new Uint8Array(actual.bytes), expected);
        if (differences.length > 0) {
          const parts = differences.map(({ part }) => part);
          const exemptions: string[] = [];
          if (staleMainCommentMarkers(new Uint8Array(actual.bytes), expected, parts, saveCommentIds, exemptions)) {
            staleMainSeeds.add(seed);
            staleMainSaves += 1;
            log.push(...exemptions);
            console.log(exemptions.map((line) => `seed=${seed} save=${saveNumber} ${line}`).join('\n'));
          } else {
            mismatchingSeeds.add(seed);
            failures.push(
              `seed=${seed} topology=${topology} save=${saveNumber}\n` +
              `${log.join('\n')}\n${differences.map(({ text }) => text).join('\n')}`
            );
          }
        }
        if (!current.peer && actual.updates.length !== 0) {
          throw new Error('Viewer save returned peer updates without a peer');
        }
      };
      const compareSave = async () => {
        if (!current.peer && current.hydration) {
          log.push('editor save requested during hydration; wait for editing copy');
          await current.hydration;
        }
        await flushPeer(current);
        const dirty = JSON.stringify([...current.main.dirtyStories].sort());
        const comments = hostComments(current.main.host).map(({ id }) => id);
        log.push(`save ${++saveNumber} dirty=${dirty} comments=${JSON.stringify(comments)}`);
        compareSaved(await saveWorkerArm(current, () => current.main.save()), comments.map(String));
      };
      if (topology !== 'C/hydrate') {
        await compareSave();
        await compareSave();
      }
      await bootstrap(current);
      const planned = operations(topology, random, hydrationMode);
      const automatic = planned.map((operation, index) => operation.autoSave === false ? -1 : index)
        .filter((index) => index >= 0);
      const forcedSave = random.pick(automatic);
      for (let index = 0; index < planned.length; index += 1) {
        log.push(`op ${index + 1}/${planned.length}: ${JSON.stringify(planned[index])}`);
        const operation = planned[index]!;
        if (operation.action === 'hydrate') {
          log.push(`${hydrationMode === 'viewer' ? 'switch viewer to editing' : 'load editable copy'} (same session)`);
          await hydrateArms(current);
        } else if (operation.action === 'hydrateSave') {
          const posted = current.worker.requests.length;
          current.worker.hold();
          const ready = hydrateArms(current);
          await applyOperation(current, { action: 'hostChange', story: 'body' }, random);
          const saving = compareSave();
          try {
            await Promise.resolve();
            expect(current.peer).toBeUndefined();
            expect(current.worker.requests.slice(posted)).toEqual(['encodeState']);
          } finally {
            current.worker.release();
            await Promise.all([ready, saving]);
          }
        } else if (operation.action === 'save' || operation.action === 'overlapSave') {
          if (current.peer || topology !== 'C/hydrate') focus(current, operation.story, random);
          if (operation.action === 'overlapSave' && current.peer) {
            await flushPeer(current);
            const dirty = JSON.stringify([...current.main.dirtyStories].sort());
            const comments = hostComments(current.main.host).map(({ id }) => id);
            log.push(`overlap save dirty=${dirty} comments=${JSON.stringify(comments)}`);
            const oracle = () => current.main.save();
            const saves = await Promise.all([saveWorkerArm(current, oracle), saveWorkerArm(current, oracle)]);
            for (const [part, saved] of saves.entries()) {
              log.push(`overlap save ${++saveNumber} (${part + 1}/2)`);
              compareSaved(saved, comments.map(String));
            }
          } else {
            await compareSave();
          }
        } else {
          await applyOperation(current, operation, random);
        }
        if (current.peer && random.int(4) === 0) await flushPeer(current);
        if (operation.autoSave !== false && (index === forcedSave || random.int(6) === 0)) {
          await compareSave();
          if (random.int(4) === 0) await compareSave();
        }
      }
      await compareSave();
      await compareSave();
      if (topology === 'B/viewer' && current.worker.requests.some(
        (type) => type === 'applyUpdate' || type === 'encodeState'
      )) {
        throw new Error('The viewer arm received peer state');
      }
      if (topology === 'C/hydrate' && (!current.peer || !current.hydration)) {
        throw new Error('The hydration arm never loaded its editing copy');
      }
    } catch (error) {
      erroredSeeds += 1;
      failures.push(
        `seed=${seed} topology=${topology}\n${log.join('\n')}\n` +
        (error instanceof Error ? error.stack : String(error))
      );
    } finally {
      arms?.unsubscribe?.();
      handbackRemoteUpdates += arms?.handbackRemoteUpdates ?? 0;
      otherRemoteUpdates += arms?.otherRemoteUpdates ?? 0;
      arms?.client.destroy();
      arms?.replica?.destroy();
      arms?.main.session.destroy();
    }
  }
  console.log(
    `resident save differential: seeds run=${seedsRun}, saves compared=${savesCompared}, ` +
    `mismatching seeds=${mismatchingSeeds.size}, errored seeds=${erroredSeeds}, ` +
    `stale main-thread comment markers: ${staleMainSaves} saves in ${staleMainSeeds.size} seeds ` +
    `[${[...staleMainSeeds].join(',')}], ` +
    `handback remote updates=${handbackRemoteUpdates}, other remote updates=${otherRemoteUpdates}`
  );
  console.log(
    `resident save hydration: seeds run=${hydrationSeeds.viewer + hydrationSeeds.unloaded + hydrationSeeds.loading}, ` +
    `viewer=${hydrationSeeds.viewer}, unloaded=${hydrationSeeds.unloaded}, loading=${hydrationSeeds.loading}`
  );
  if (failures.length > 0) console.error(failures.join('\n\n'));
  expect(failures.length).toBe(0);
  if (editorSeeds > 0) expect(handbackRemoteUpdates).toBeGreaterThan(0);
}, Math.max(120_000, SEEDS * OPS * 30));
