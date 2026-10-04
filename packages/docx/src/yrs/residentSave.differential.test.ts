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

async function saveWorkerArm(arms: Arms) {
  const dirty = arms.peer ? arms.editorStories.captureWorkerSave() : undefined;
  arms.log.push(`worker save stories=${dirty ? JSON.stringify([...dirty.stories].sort()) : 'resident'}`);
  const saved = await arms.client.save({
    comments: hostComments(arms.workerHost),
    host: hostSaveMetadata(arms.workerHost),
    ...(arms.peer && dirty ? { stateVector: arms.peer.encodeStateVector(), stories: dirty.stories } : {}),
  });
  dirty?.clear();
  return saved;
}

const DATE = '2026-10-02T12:00:00Z';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const STORIES = ['body', 'hf:rIdH1', 'fn:1'] as const;
type Story = (typeof STORIES)[number];
type Topology = 'A/editor' | 'B/viewer';
type Action =
  | 'insert' | 'delete' | 'addComment' | 'reply' | 'deleteComment' | 'workerDeleteComment'
  | 'proposal' | 'decide' | 'withdraw' | 'flush' | 'project' | 'undo' | 'redo' | 'save';
interface Operation {
  action: Action;
  story: Story;
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
  peer?: YrsSession;
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
  try {
    const { hostJson } = await client.open(source, { generation });
    session = await createYrsSession({ clientId: clientId + (topology === 'A/editor' ? 1 : 0) });
    const { document: host } = session.openDocx(source.slice(), topology === 'B/viewer', { generation });
    if (topology === 'A/editor') {
      session.loadState(await client.encodeState());
      session.beginUndoCapture();
    }
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
      adoptingWorkerSaveUpdates: false, handbackRemoteUpdates: 0, otherRemoteUpdates: 0,
      ...(topology === 'A/editor' ? { peer: session } : {}),
      commentStories: new Map<number, Story>([[1, 'body'], [2, 'hf:rIdH1']]),
      nextComment: 3, nextProposal: 1,
    };
    if (topology === 'A/editor') {
      const session = arms.main.session;
      // PagedEditor remote listener -> publishDirectInput(undefined).
      arms.unsubscribe = session.onUpdate((_update, origin) => {
        if (origin !== 'remote') return;
        if (arms.adoptingWorkerSaveUpdates) arms.handbackRemoteUpdates += 1;
        else arms.otherRemoteUpdates += 1;
        if (!session.hasStory('body')) return;
        editorStories.add(dirtyProjectionStory(session.selection()?.head.story ?? 'body'));
      });
    }
    return arms;
  } catch (error) {
    client.destroy();
    session?.destroy();
    throw error;
  }
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
  arms.mirror = reply.mirror;
  return reply;
}

function target(arms: Arms, story: Story, random: Random) {
  const paragraphs = arms.main.session.paragraphs(story);
  const editable = story === 'body' ? paragraphs.filter(({ paraId }) => paraId !== '0000B004') : paragraphs;
  const paragraph = random.pick(editable);
  const offset = random.int(paragraph.text.length + 1);
  const at = { story, paraId: paragraph.paraId, offset };
  return { at, paragraph };
}

function focus(arms: Arms, story: Story, random: Random) {
  const chosen = target(arms, story, random);
  arms.main.session.setSelection(chosen.at);
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
    arms.peer.applyRawOps(story, [{
      op: 'setComment', id: String(comment.id), ranges: [[start, start + 1]],
      author: comment.author, date: comment.date, body: comment.content,
    }]);
    arms.main.publishDirectInput();
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
    await textProposal(arms, random.pick(STORIES), random, ' proposed');
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
  const session = arms.main.session;
  switch (action) {
    case 'insert':
    case 'delete': {
      const text = random.pick([' x', 'Y', '  z ', 'q&']);
      if (!arms.peer) {
        await textProposal(arms, story, random, text, action === 'delete');
        return;
      }
      const { at, paragraph } = focus(arms, story, random);
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
      arms.main.publishDirectInput();
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
      if (!arms.peer || action === 'deleteComment') {
        try {
          session.applyRawOps('body', [{ op: 'removeComment', id: String(id) }]);
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
      const changed = action === 'undo' ? session.undo() : session.redo();
      const stories = changed ? session.historyStories() : [];
      arms.log.push(`${action} ${changed} ${JSON.stringify(stories)}`);
      if (changed) arms.main.publishDirectInput(stories);
      return;
    }
  }
}

function operations(topology: Topology, random: Random): Operation[] {
  const group = (story: Story, actions: Action[]): Operation[] =>
    actions.map((action) => ({ action, story }));
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
    group(random.pick(STORIES), ['save', 'save'])
  );
  const required = random.shuffle(groups).flat();
  const weighted: Action[] = ['insert', 'insert', 'delete', 'delete', 'addComment', 'addComment',
    'reply', 'reply', 'deleteComment', 'workerDeleteComment', 'proposal', 'decide', 'withdraw', 'flush', 'save',
    ...(topology === 'A/editor' ? ['project', 'undo', 'redo'] as const : [])];
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

function danglingRangeMarkers(parts: Record<string, Uint8Array>, part: string): string[] {
  const xml = parts[part] ? new TextDecoder().decode(parts[part]) : '';
  const ids = liveCommentIds(parts);
  return [...xml.matchAll(/<w:commentRange(?:Start|End)\b[^>]*\bw:id="([^"]+)"/g)]
    .map((match) => match[1]!)
    .filter((id) => !ids.has(id));
}

const COMMENT_RANGE_MARKER = /<w:commentRange(?:Start|End)\b[^<>]*\bw:id="([^"]+)"[^<>]*\/>/g;
const COMMENT_REFERENCE_RUN =
  /<w:r\b[^<>]*>\s*(?:<w:rPr\b[^<>]*>(?:(?!<\/?w:r\b|<\/w:rPr>)[\s\S])*<\/w:rPr>\s*)?<w:commentReference\b[^<>]*\bw:id="([^"]+)"[^<>]*\/>\s*<\/w:r>/g;

function withoutStaleCommentMarkers(xml: string, liveIds: ReadonlySet<string>): string {
  return [COMMENT_REFERENCE_RUN, COMMENT_RANGE_MARKER].reduce((text, pattern) =>
    text.replace(pattern, (marker, id: string) => liveIds.has(id) ? marker : ''), xml
  );
}

// Deferred: after a host projection, the main-thread save keeps the range markers of a comment
// deleted outside the selection's story; the worker writes the story as it is.
function staleMainCommentMarkers(
  worker: Uint8Array, main: Uint8Array, parts: readonly string[], log: string[]
): boolean {
  const actual = unzipContainer(worker);
  const expected = unzipContainer(main);
  const liveIds = liveCommentIds(expected);
  const matches = parts.length > 0 && parts.every((part) =>
    /^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(part) &&
    danglingRangeMarkers(actual, part).length === 0 &&
    danglingRangeMarkers(expected, part).length > 0 &&
    actual[part] !== undefined && expected[part] !== undefined &&
    withoutStaleCommentMarkers(new TextDecoder().decode(expected[part]), liveIds) ===
      new TextDecoder().decode(actual[part])
  );
  if (matches) {
    for (const part of parts) {
      const xml = new TextDecoder().decode(expected[part]);
      const ids = [COMMENT_RANGE_MARKER, COMMENT_REFERENCE_RUN].flatMap((pattern) =>
        [...xml.matchAll(pattern)].map((match) => match[1]!).filter((id) => !liveIds.has(id))
      );
      log.push(`stale main-thread comment markers in ${part}: ids=${JSON.stringify([...new Set(ids)])}`);
    }
  }
  return matches;
}

test('the stale comment marker exemption removes only dangling markers', () => {
  const worker = '<w:p><w:r><w:t>text</w:t></w:r></w:p>';
  const range = '<w:commentRangeStart w:id="7"/><w:commentRangeEnd w:id="7"/>';
  const reference = '<w:r><w:commentReference w:id="7"/></w:r>';
  const formattedReference = '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' +
    '<w:commentReference w:id="7"/></w:r>';
  const main = worker.replace('</w:p>', `${range}${reference}${formattedReference}</w:p>`);
  const liveIds = new Set<string>();
  expect(withoutStaleCommentMarkers(main, liveIds)).toBe(worker);
  expect(withoutStaleCommentMarkers(main.replace('text', 'text!'), liveIds)).not.toBe(worker);
  expect(withoutStaleCommentMarkers(main.replace('</w:p>', '<x:foreign/></w:p>'), liveIds))
    .not.toBe(worker);
  expect(withoutStaleCommentMarkers(main, new Set(['7']))).toBe(main);
  const mixed = '<w:r><w:commentReference w:id="7"/><w:t>keep</w:t></w:r>';
  expect(withoutStaleCommentMarkers(mixed, liveIds)).toBe(mixed);
  const foreign = reference.replace('</w:r>', '<x:foreign/></w:r>');
  expect(withoutStaleCommentMarkers(foreign, liveIds)).toBe(foreign);
  const mixedMarkers = `<w:r>${range}<w:commentReference w:id="7"/></w:r>`;
  expect(withoutStaleCommentMarkers(mixedMarkers, liveIds)).toBe(reference);
  const bare = '<w:commentReference w:id="7"/>';
  expect(withoutStaleCommentMarkers(bare, liveIds)).toBe(bare);
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
  for (let seed = 1; seed <= SEEDS; seed += 1) {
    const random = new Random(seed);
    const topology: Topology = random.int(2) === 0 ? 'A/editor' : 'B/viewer';
    const log: string[] = [`open seed=${seed} topology=${topology}`];
    let arms: Arms | undefined;
    seedsRun += 1;
    try {
      arms = await openArms(seed, topology, log);
      const current = arms;
      let saveNumber = 0;
      const compareSave = async () => {
        await flushPeer(current);
        const dirty = JSON.stringify([...current.main.dirtyStories].sort());
        const comments = JSON.stringify(hostComments(current.main.host).map(({ id }) => id));
        log.push(`save ${++saveNumber} dirty=${dirty} comments=${comments}`);
        const actual = await saveWorkerArm(current);
        const expected = await current.main.save();
        savesCompared += 1;
        const differences = partDifferences(new Uint8Array(actual.bytes), expected);
        if (differences.length > 0) {
          const parts = differences.map(({ part }) => part);
          const exemptions: string[] = [];
          if (staleMainCommentMarkers(new Uint8Array(actual.bytes), expected, parts, exemptions)) {
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
        if (current.peer) {
          if (actual.updates.length !== 1) throw new Error('Editor save did not return exactly one diff');
          current.adoptingWorkerSaveUpdates = true;
          try {
            current.editorStories.adoptWorkerSaveUpdates(() => {
              for (const update of actual.updates) current.peer!.applyUpdate(update);
            });
          } finally {
            current.adoptingWorkerSaveUpdates = false;
          }
        } else if (actual.updates.length !== 0) {
          throw new Error('Viewer save returned peer updates without a peer');
        }
      };
      await compareSave();
      await compareSave();
      await bootstrap(current);
      const planned = operations(topology, random);
      const forcedSave = random.int(planned.length);
      for (let index = 0; index < planned.length; index += 1) {
        log.push(`op ${index + 1}/${planned.length}: ${JSON.stringify(planned[index])}`);
        const operation = planned[index]!;
        if (operation.action === 'save') {
          focus(current, operation.story, random);
          await compareSave();
        } else {
          await applyOperation(current, operation, random);
        }
        if (current.peer && random.int(4) === 0) await flushPeer(current);
        if (index === forcedSave || random.int(6) === 0) {
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
  if (failures.length > 0) console.error(failures.join('\n\n'));
  expect(failures.length).toBe(0);
}, Math.max(120_000, SEEDS * OPS * 30));
