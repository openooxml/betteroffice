import { afterEach, beforeAll, describe, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsSession,
  computeProposalGeometryMirror,
  proposalSetIdentity,
  type DocxEditRequest,
  type ResidentProposalReply,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type { PluginInvocation } from '../../../../shared/plugin-host/runtime';
import { UNAVAILABLE_DOCX_COMMANDS } from '../commands/createDocxCommandStore';
import * as editorBatches from '../components/DocxEditor/editorBatches';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';
import { stampRevisionPreviewKey, stampSourceVersion } from '../components/DocxEditor/internals/layoutProvenance';
import * as workerOpenReplica from '../components/DocxEditor/internals/workerOpenReplica';
import * as workerProposals from '../components/DocxEditor/internals/workerProposalAuthority';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import { createDocxPluginHost } from './createDocxPluginHost';
import { createPluginClients, resolveParagraph } from './createPluginClients';
import { defineDocxPlugin } from './defineDocxPlugin';
import { currentPreviewKey } from './proposalPreview';
import type { DocxPluginContext, DocxPluginGrant, DocxPluginSnapshot } from './types';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const WORD = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (id: string, text: string) => `<w:p w14:paraId="${id}">${run(text)}</w:p>`;

function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WORD}.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${WORD}.header+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/_rels/document.xml.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document ${NS}><w:body>${paragraph('00000001', 'Alpha')}${paragraph(
        '00000002',
        'Tail'
      )}<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${paragraph(
        '0000A001',
        'Locked'
      )}</w:sdtContent></w:sdt><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:body></w:document>`
    )
  );
  parts.set('word/header1.xml', toBytes(`<w:hdr ${NS}>${paragraph('0000E001', 'Header')}</w:hdr>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
const sessions: YrsSession[] = [];
const restoreClocks: Array<() => void> = [];
const restoreWorkers: Array<() => void> = [];
afterEach(() => {
  for (const restore of restoreWorkers.splice(0)) restore();
  for (const restore of restoreClocks.splice(0)) restore();
  for (const session of sessions.splice(0)) session.destroy();
});

function navigationClock() {
  let now = 0;
  let nextTimer = 0;
  const timers = new Map<number, { at: number; callback: () => void }>();
  const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(
    ((callback: () => void, delay = 0) => {
      const id = ++nextTimer;
      timers.set(id, { at: now + delay, callback });
      return id;
    }) as unknown as typeof setTimeout
  );
  const clear = spyOn(globalThis, 'clearTimeout').mockImplementation((id) => {
    timers.delete(id as unknown as number);
  });
  restoreClocks.push(() => {
    timeout.mockRestore();
    clear.mockRestore();
  });
  return {
    timers,
    advance(ms: number) {
      now += ms;
      for (const [id, timer] of timers) {
        if (timer.at > now) continue;
        timers.delete(id);
        timer.callback();
      }
    },
  };
}

async function setup(
  options: { grant?: DocxPluginGrant; flush?: () => void | Promise<void> } = {}
) {
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(fixture(), true);
  const events: string[] = [];
  const editor = {
    getYrsSession: () => session,
    flushPendingInput: async () => {
      events.push('flush');
      await options.flush?.();
    },
    syncYrsInputState: (docChanged: boolean, stories?: readonly string[]) => {
      events.push(`sync:${docChanged}:${stories?.join(',') ?? '*'}`);
      return true;
    },
    revealDisplayPosition: (position: number) => {
      events.push(`scroll:${position}`);
      return state.anchorReady ? state.reveal : 'unsupported';
    },
    focus: () => events.push('focus'),
    setSelection: (anchor: number) => events.push(`select:${anchor}`),
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor as PagedEditorRef | null };
  const state = {
    mode: 'editing' as EditorMode,
    grant: options.grant ?? ({ document: 'write', editBatches: true } as DocxPluginGrant),
    layoutReady: true,
    partial: false,
    anchorReady: true,
    layoutFailed: false,
    queryState: 'ready' as 'loading' | 'ready' | 'error',
    reveal: 'scrolled' as 'scrolled' | 'layout-unavailable' | 'unsupported',
    ended: null as 'plugin-unavailable' | 'document-replaced' | null,
    viewer: false,
  };
  const controller = new AbortController();
  const lifetimeController = new AbortController();
  const invocation: PluginInvocation<DocxPluginSnapshot> = {
    pluginId: 'acme.review',
    activation: {},
    snapshot: {} as DocxPluginSnapshot,
    signal: controller.signal,
    lifetimeSignal: lifetimeController.signal,
    state: () => null,
    setState: () => false,
    onCleanup: () => {},
    run: async () => {},
    commit: (write) => write(),
    refusal: () => state.ended ?? (controller.signal.aborted ? 'aborted' : null),
  };
  const queries = {
    sourceState: () => ({ status: state.queryState }),
    anchorRect: () =>
      state.anchorReady ? { pageIndex: 0, x: 0, y: 0, width: 1, height: 1 } : null,
  } as unknown as DisplayListQueries;
  stampSourceVersion(queries, session.version());
  const layoutListeners = new Set<() => void>();
  let layoutStarted!: () => void;
  const waiting = new Promise<void>((resolve) => (layoutStarted = resolve));
  const access = {
    pagedEditorRef,
    writeMode: () => state.mode,
    viewer: () => state.viewer,
    commands: () => null,
    layout: () => ({
      queries: state.layoutReady ? queries : null,
      complete: !state.partial,
      failed: state.layoutFailed,
    }),
    subscribeLayout(listener: () => void) {
      layoutListeners.add(listener);
      const unsubscribe = session.onUpdate(listener);
      layoutStarted();
      return () => {
        layoutListeners.delete(listener);
        unsubscribe();
      };
    },
  };
  const createClients = () =>
    createPluginClients(invocation, access, () => state.grant, UNAVAILABLE_DOCX_COMMANDS);
  return {
    session,
    events,
    pagedEditorRef,
    editor,
    state,
    controller,
    lifetimeController,
    queries,
    access,
    clients: createClients(),
    createClients,
    waiting,
    layoutListeners,
    publishLayout: () => {
      for (const listener of layoutListeners) listener();
    },
  };
}

function replace(version: string, paraId: string, text: string, story = 'body'): DocxEditRequest {
  return {
    expectVersion: version,
    steps: [{ op: 'replaceText', target: { kind: 'paragraph', story, paraId }, text }],
  };
}

function texts(session: YrsSession): string[] {
  return session.paragraphs('body').map((candidate) => candidate.text);
}

function routeWorker(env: Awaited<ReturnType<typeof setup>>) {
  const authority: Pick<
    workerProposals.WorkerProposalAuthority,
    'initialized' | 'navigationTarget' | 'readParagraphs' | 'findText'
  > = {
    initialized: true,
    async navigationTarget(story: string, paraId: string) {
      return {
        version: env.session.version(),
        target: { loc: { story, paraId, offset: 0 }, position: 42 },
      };
    },
    async findText() {
      return { ok: true as const, version: env.session.version(), matches: [], truncated: false };
    },
    async readParagraphs(request) {
      return {
        ok: true as const,
        version: env.session.version(),
        view: request.view,
        paragraphs: [{ story: 'body', paraId: 'worker', text: 'Worker text', atoms: [] }],
      };
    },
  };
  const lookup = workerProposals.workerProposalAuthority;
  const routing = spyOn(workerProposals, 'workerProposalAuthority').mockImplementation((session) =>
    session === env.session
      ? (authority as workerProposals.WorkerProposalAuthority)
      : lookup(session)
  );
  const navigation = spyOn(authority, 'navigationTarget');
  const read = spyOn(authority, 'readParagraphs');
  const findText = spyOn(authority, 'findText');
  const flush = spyOn(editorBatches, 'flushEditorInput');
  const replica = spyOn(workerOpenReplica, 'requestWorkerOpenReplica');
  restoreWorkers.push(() => {
    replica.mockRestore();
    flush.mockRestore();
    read.mockRestore();
    findText.mockRestore();
    navigation.mockRestore();
    routing.mockRestore();
  });
  return { authority, navigation, read, findText, flush, replica, routing };
}

describe('plugin edit client', () => {
  test('applies with the expected version and refuses one that flushed typing made stale', async () => {
    let session!: YrsSession;
    const typing = { active: false };
    const env = await setup({
      flush: () => {
        if (typing.active)
          session.insertText({ story: 'body', paraId: '00000002', offset: 4 }, ' typed');
      },
    });
    session = env.session;
    const version = session.version();
    typing.active = true;
    expect(await env.clients.edits!.applyEdits(replace(version, '00000001', 'Beta'))).toMatchObject(
      {
        ok: false,
        failure: { code: 'stale-version' },
      }
    );
    expect(texts(session)[1]).toBe('Tail typed');
    typing.active = false;
    const applied = await env.clients.edits!.applyEdits(
      replace(session.version(), '00000001', 'Beta')
    );
    expect(applied).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
    expect(env.events.filter((event) => event.startsWith('sync'))).toEqual(['sync:true:body']);
  });

  test('revocation, viewing mode, cancellation and replacement during the flush refuse', async () => {
    const cases: Array<[string, (env: Awaited<ReturnType<typeof setup>>) => void, unknown]> = [
      ['revoked', (env) => (env.state.grant = {}), { failure: { code: 'permission-denied' } }],
      ['viewing', (env) => (env.state.mode = 'viewing'), { failure: { code: 'read-only' } }],
      ['aborted', (env) => env.controller.abort(), { failure: { code: 'aborted' } }],
      [
        'replaced',
        (env) =>
          (env.pagedEditorRef.current = {
            ...env.editor,
            getYrsSession: () => null,
          } as unknown as PagedEditorRef),
        { failure: { code: 'document-replaced' } },
      ],
    ];
    for (const [, during, expected] of cases) {
      let env!: Awaited<ReturnType<typeof setup>>;
      env = await setup({ flush: () => during(env) });
      const version = env.session.version();
      expect(
        await env.clients.edits!.applyEdits(replace(version, '00000001', 'Beta'))
      ).toMatchObject({
        ok: false,
        ...(expected as object),
      });
      expect(env.session.version()).toBe(version);
    }
  });

  test('document policy refusals come back from the engine unchanged', async () => {
    const env = await setup();
    expect(
      await env.clients.edits!.applyEdits(
        replace(env.session.version(), '0000A001', 'x', 'body:sdt0')
      )
    ).toMatchObject({ ok: false, failure: { code: 'locked-target' } });
  });

  test('untracked history needs its own grant; no grant means no edit client', async () => {
    const env = await setup();
    const request = {
      ...replace(env.session.version(), '00000001', 'Beta'),
      history: 'none' as const,
    };
    expect(await env.clients.edits!.applyEdits(request)).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });
    env.state.grant = { document: 'write', editBatches: true, untrackedHistory: true };
    expect(await env.clients.edits!.applyEdits(request)).toMatchObject({ ok: true, applied: true });
    expect(env.session.canUndo()).toBe(false);
    expect((await setup({ grant: { document: 'write' } })).clients.edits).toBeNull();
  });
});

describe('plugin read and navigation clients', () => {
  test('worker navigation waits for the current preview at the same document version', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    const proposed = env.session.proposeChanges({
      expectVersion: env.session.version(),
      proposals: [{
        id: 'preview-proposal',
        paragraph: { kind: 'session', sessionId: env.session.paragraphIdentities().sessionId, story: 'body', paraId: '00000001' },
        suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
        op: 'insertText', at: 'start', text: 'Added ',
      }],
    });
    expect(proposed).toMatchObject({ ok: true });
    const snapshot = env.session.getProposals();
    stampSourceVersion(env.queries, snapshot.version);
    stampRevisionPreviewKey(env.queries, currentPreviewKey(env.session));
    expect(env.session.setProposalStates({
      expectVersion: snapshot.version,
      expectPreviewVersion: snapshot.previewVersion,
      changes: [{ id: 'preview-proposal', state: 'rejected' }],
    })).toMatchObject({ ok: true });
    expect(env.session.version()).toBe(snapshot.version);
    expect(currentPreviewKey(env.session)).not.toBe('');
    const scroll = env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: '00000002' }, { expectVersion: snapshot.version }
    );
    await env.waiting;
    expect(env.layoutListeners.size).toBe(1);
    env.publishLayout();
    expect(env.events).toEqual([]);
    stampRevisionPreviewKey(env.queries, currentPreviewKey(env.session));
    env.publishLayout();
    expect(await scroll).toEqual({ ok: true });
    expect(worker.navigation).toHaveBeenCalledTimes(1);
    expect(env.events).toEqual(['scroll:42']);
    expect(worker.replica).not.toHaveBeenCalled();
  });

  test('worker navigation resolves once and reveals without flushing or requesting the replica', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    const target = { story: 'body', paraId: '00000002' };
    const selection = JSON.stringify(env.session.selection());
    expect(
      await env.clients.navigation.scrollToParagraph(target, {
        expectVersion: env.session.version(),
      })
    ).toEqual({ ok: true });
    expect(worker.navigation).toHaveBeenCalledTimes(1);
    expect(worker.navigation.mock.calls[0]?.slice(0, 2)).toEqual(['body', '00000002']);
    expect(worker.flush).not.toHaveBeenCalled();
    expect(worker.replica).not.toHaveBeenCalled();
    expect(env.events).toEqual(['scroll:42']);
    expect(JSON.stringify(env.session.selection())).toBe(selection);
  });

  for (const toggleOnly of [true, false]) {
    test(`worker navigation uses ${toggleOnly ? 'zero reads after a toggle' : 'one read after a version-changing proposal'} and matches repeated resolution`, async () => {
      const env = await setup();
      const paragraph = env.session.paragraphIdentities().paragraphs[1]!.session!;
      const target = { story: paragraph.story, paraId: paragraph.paraId };
      const before = env.session.version();
      expect(env.session.proposeChanges({
        expectVersion: env.session.version(),
        proposals: [{
          id: 'jump', paragraph, op: 'insertText', at: 'start', text: 'Proposed ',
          suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
        }],
      }).ok).toBe(true);
      expect(env.session.version()).not.toBe(before);
      const proposed = env.session.getProposals();
      const proposedGeometry = computeProposalGeometryMirror(env.session, proposed, false);
      expect(proposedGeometry.navigationTargets).toBeUndefined();
      if (toggleOnly) {
        expect(env.session.setProposalStates({
          expectVersion: proposed.version,
          expectPreviewVersion: proposed.previewVersion,
          changes: [{ id: 'jump', state: 'rejected' }],
        }).ok).toBe(true);
        expect(env.session.version()).toBe(proposed.version);
        expect(env.session.getProposals().previewVersion).toBe(proposed.previewVersion + 1);
      }
      const first = resolveParagraph(env.session, target);
      const second = resolveParagraph(env.session, target);
      expect(first).toEqual(second);
      if (typeof second === 'string') throw new Error('expected a navigation target');
      const version = 'worker-1';
      const proposals = env.session.getProposals();
      const geometry = toggleOnly ? computeProposalGeometryMirror(env.session, proposals) : proposedGeometry;
      if (toggleOnly) expect(geometry.navigationTargets?.jump).toEqual(second);
      const snapshot: ResidentProposalReply = {
        mirror: {
          version,
          proposals: {
            previewVersion: proposals.previewVersion,
            entries: proposals.proposals.map((record) => ({
              record, key: 'jump', suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
            })),
          },
        },
        geometry: { ...geometry, version },
        changedStories: toggleOnly ? [] : ['body'], updates: [], stateVector: new Uint8Array(),
      };
      const documentRead = mock(async () => ({ version, value: second }) as never);
      const replica = workerOpenReplica.deferWorkerOpenReplica(env.session, async () => {
        throw new Error('navigation must not hydrate the replica');
      }, () => { throw new Error('unexpected fallback'); }, () => {});
      restoreWorkers.push(() => replica.cancel());
      const authority = workerProposals.registerWorkerProposalAuthority(env.session, {
        proposal: async () => snapshot,
        documentRead,
        handOver: async () => { throw new Error('unexpected hand-over'); },
      }, {
        laidOut: async () => {}, current: () => true, relayout: () => {}, contentChanged: () => {},
        adopted: () => {}, handedOver: () => {},
      });
      await authority.initialize();
      stampSourceVersion(env.queries, version);
      stampRevisionPreviewKey(env.queries, currentPreviewKey(env.session));
      const count = spyOn(env.session, 'paragraphIdCount').mockImplementation(() => {
        throw new Error('navigation must not read the unhydrated replica');
      });
      const flush = spyOn(editorBatches, 'flushEditorInput');
      restoreWorkers.push(() => { count.mockRestore(); flush.mockRestore(); });
      env.state.layoutReady = false;
      const scroll = env.clients.navigation.scrollToParagraph(target, { expectVersion: version });
      await env.waiting;
      expect(documentRead).toHaveBeenCalledTimes(toggleOnly ? 0 : 1);
      expect(env.events).toEqual([]);
      env.state.layoutReady = true;
      env.publishLayout();
      expect(await scroll).toEqual({ ok: true });
      expect(documentRead).toHaveBeenCalledTimes(toggleOnly ? 0 : 1);
      expect(env.events).toEqual([`scroll:${second.position}`]);
      expect(flush).not.toHaveBeenCalled();
      expect(replica.started).toBe(false);
      expect(env.layoutListeners.size).toBe(0);
    });
  }

  test('worker navigation checks the session and reply versions around each resolution', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    const target = { story: 'body', paraId: '00000002' };
    const version = env.session.version();
    const stale = {
      ok: false,
      failure: { code: 'stale-version', message: 'The document changed after that version' },
    } as const;
    expect(
      await env.clients.navigation.scrollToParagraph(target, { expectVersion: 'older' })
    ).toEqual(stale);
    expect(worker.navigation).not.toHaveBeenCalled();
    worker.navigation.mockResolvedValueOnce({
      version: 'older',
      target: { loc: { ...target, offset: 0 }, position: 42 },
    });
    expect(
      await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })
    ).toEqual(stale);
    let release!: () => void;
    const waiting = new Promise<void>((resolve) => (release = resolve));
    worker.navigation.mockImplementationOnce(async () => {
      await waiting;
      return { version, target: { loc: { ...target, offset: 0 }, position: 42 } };
    });
    const scroll = env.clients.navigation.scrollToParagraph(target, { expectVersion: version });
    env.session.insertText({ story: 'body', paraId: '00000001', offset: 5 }, '!');
    release();
    expect(await scroll).toEqual(stale);
    expect(worker.flush).not.toHaveBeenCalled();
    expect(worker.replica).not.toHaveBeenCalled();
    expect(env.events).toEqual([]);
  });

  test('worker navigation rechecks the version after layout settles without reading again', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    const target = { story: 'body', paraId: '00000002' };
    const version = env.session.version();
    const layout = spyOn(env.access, 'layout').mockImplementationOnce(() => {
      queueMicrotask(() => {
        env.session.insertText({ story: 'body', paraId: '00000001', offset: 5 }, '!');
      });
      return { queries: env.queries, complete: true, failed: false };
    });
    restoreWorkers.push(() => layout.mockRestore());
    expect(
      await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })
    ).toMatchObject({
      ok: false,
      failure: { code: 'stale-version' },
    });
    expect(worker.navigation).toHaveBeenCalledTimes(1);
    expect(env.events).toEqual([]);
  });

  test('worker navigation preserves target failures and waits for the rendered layout', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    const version = env.session.version();
    const target = { story: 'body', paraId: '00000002' };
    worker.navigation.mockResolvedValueOnce({ version, target: 'missing-target' });
    expect(
      await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })
    ).toEqual({
      ok: false,
      failure: { code: 'missing-target', message: 'The paragraph cannot be shown (missing-target)' },
    });
    env.state.layoutReady = false;
    const scroll = env.clients.navigation.scrollToParagraph(target, { expectVersion: version });
    await env.waiting;
    expect(env.events).toEqual([]);
    env.state.layoutReady = true;
    env.publishLayout();
    expect(await scroll).toEqual({ ok: true });
    expect(env.events).toEqual(['scroll:42']);
    expect(worker.flush).not.toHaveBeenCalled();
    expect(worker.replica).not.toHaveBeenCalled();
  });

  test('worker navigation with focus waits for the replica and checks replacement', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    const target = { story: 'body', paraId: '00000002' };
    const version = env.session.version();
    let release!: () => void;
    let requested!: () => void;
    const ready = new Promise<void>((resolve) => (release = resolve));
    const waiting = new Promise<void>((resolve) => (requested = resolve));
    worker.replica.mockImplementation(() => {
      requested();
      return ready;
    });
    const scroll = env.clients.navigation.scrollToParagraph(target, {
      expectVersion: version,
      focus: true,
    });
    await waiting;
    expect(worker.replica).toHaveBeenCalledWith(env.session);
    expect(env.events).toEqual([]);
    release();
    expect(await scroll).toEqual({ ok: true });
    expect(env.events).toEqual(['scroll:42', 'sync:false:*', 'focus']);
    expect(env.session.selection()?.head).toMatchObject({ ...target, offset: 0 });
    worker.replica.mockImplementation(async () => {
      env.pagedEditorRef.current = null;
    });
    expect(
      await env.clients.navigation.scrollToParagraph(target, {
        expectVersion: version,
        focus: true,
      })
    ).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
    expect(env.events).toEqual(['scroll:42', 'sync:false:*', 'focus']);
    expect(worker.flush).not.toHaveBeenCalled();
  });

  test('viewer navigation with focus selects in the viewer input without the replica', async () => {
    const env = await setup();
    env.state.viewer = true;
    const worker = routeWorker(env);
    const before = env.session.selection();
    expect(
      await env.clients.navigation.scrollToParagraph(
        { story: 'body', paraId: '00000002' },
        { expectVersion: env.session.version(), focus: true }
      )
    ).toEqual({ ok: true });
    expect(env.events).toEqual(['scroll:42', 'select:42', 'focus']);
    expect(worker.replica).not.toHaveBeenCalled();
    expect(env.session.selection()).toEqual(before);
  });

  for (const during of ['target read', 'layout wait'] as const) {
    test(`navigation accepts an equivalent worker version after hand-over during ${during}`, async () => {
      const env = await setup();
      env.state.layoutReady = false;
      const version = 'worker-1';
      const snapshot: ResidentProposalReply = {
        mirror: { version, proposals: { previewVersion: 0, entries: [] } },
        geometry: {
          version, previewVersion: 0, proposals: proposalSetIdentity(env.session.getProposals()),
          targets: {}, hidden: [],
        },
        changedStories: [], updates: [], stateVector: new Uint8Array(),
      };
      let release!: () => void;
      const transfer = new Promise<void>((resolve) => { release = resolve; });
      workerOpenReplica.deferWorkerOpenReplica(env.session, async () => {
        const handover = await workerProposals.beginWorkerProposalHandover(env.session)!;
        return () => { handover.complete(); };
      }, () => { throw new Error('unexpected fallback'); }, () => {});
      const authority = workerProposals.registerWorkerProposalAuthority(env.session, {
        proposal: async () => snapshot,
        documentRead: async () => { throw new Error('unexpected worker read'); },
        handOver: async () => {
          await transfer;
          return { version, state: new Uint8Array(), proposals: snapshot.mirror.proposals };
        },
      }, {
        laidOut: () => Promise.resolve(), current: () => true, relayout: () => {}, contentChanged: () => {},
        adopted: (token) => workerOpenReplica.adoptWorkerOpenMirrorVersion(env.session, token),
        handedOver: (token) => workerOpenReplica.adoptWorkerOpenHandoverVersion(env.session, token),
      });
      await authority.initialize();
      stampSourceVersion(env.queries, version);
      const target = { story: 'body', paraId: '00000002' };
      const navigation = spyOn(authority, 'navigationTarget');
      restoreWorkers.push(() => navigation.mockRestore());
      if (during === 'layout wait') {
        navigation.mockImplementationOnce(async () => ({
          version, target: resolveParagraph(env.session, target),
        }));
      }
      const ready = during === 'target read'
        ? workerOpenReplica.requestWorkerOpenReplica(env.session)!
        : null;
      const scroll = env.clients.navigation.scrollToParagraph(target, { expectVersion: version });
      if (during === 'layout wait') await env.waiting;
      const hydrated = ready ?? workerOpenReplica.requestWorkerOpenReplica(env.session)!;
      release();
      await hydrated;
      await env.waiting;
      expect(env.session.version()).not.toBe(version);
      expect(env.events).toEqual([]);
      env.state.layoutReady = true;
      env.publishLayout();
      expect(await scroll).toEqual({ ok: true });
      expect(env.events.filter((event) => event.startsWith('scroll'))).toHaveLength(1);
      expect(env.layoutListeners.size).toBe(0);

      env.session.insertText({ story: 'body', paraId: '00000001', offset: 5 }, '!');
      expect(await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })).toMatchObject({
        ok: false, failure: { code: 'stale-version' },
      });
    });
  }

  test('worker navigation keeps supersession and invocation aborts across awaited resolutions', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    const version = env.session.version();
    const target = { story: 'body', paraId: '00000002' };
    let release!: () => void;
    const waiting = new Promise<void>((resolve) => (release = resolve));
    worker.navigation.mockImplementationOnce(async () => {
      await waiting;
      return { version, target: { loc: { ...target, offset: 0 }, position: 99 } };
    });
    const older = env.clients.navigation.scrollToParagraph(target, { expectVersion: version });
    expect(
      await env.createClients().navigation.scrollToParagraph(target, { expectVersion: version })
    ).toEqual({ ok: true });
    release();
    expect(await older).toEqual({
      ok: false,
      failure: {
        code: 'layout-unavailable',
        message: 'A newer paragraph navigation superseded this request',
      },
    });
    expect(env.events).toEqual(['scroll:42']);
    worker.navigation.mockImplementationOnce(async () => {
      await Promise.resolve();
      env.controller.abort();
      return { version, target: { loc: { ...target, offset: 0 }, position: 99 } };
    });
    expect(
      await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })
    ).toMatchObject({ ok: false, failure: { code: 'aborted' } });
    expect(env.events).toEqual(['scroll:42']);
    expect(worker.flush).not.toHaveBeenCalled();
    expect(worker.replica).not.toHaveBeenCalled();
  });

  test('editor worker reads skip flushing while text searches retain replica admission', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    expect(await env.clients.read.version()).toEqual({ ok: true, version: env.session.version() });
    const read = await env.clients.read.readParagraphs({ view: 'accepted' });
    expect(read.ok && read.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Worker text']);
    expect(worker.read).toHaveBeenCalledTimes(1);
    expect(worker.flush).not.toHaveBeenCalled();
    expect(worker.replica).not.toHaveBeenCalled();
    expect(
      await env.clients.read.findText({
        text: 'Tail',
        within: { kind: 'story', story: 'body' },
        view: 'accepted',
      })
    ).toMatchObject({ ok: true });
    expect(worker.replica).toHaveBeenCalledWith(env.session);
    expect(worker.flush).toHaveBeenCalledTimes(1);
    expect(env.events).toEqual(['flush']);
  });

  test('worker reads use the flushed fallback and keep plugin refusals during hand-over', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    worker.read.mockImplementation((request, main) => main(request));
    const read = await env.clients.read.readParagraphs({ view: 'accepted' });
    expect(read.ok && read.paragraphs[0]?.text).toBe('Alpha');
    expect(env.events).toEqual(['flush']);
    const pending = spyOn(env.editor, 'flushPendingInput').mockImplementationOnce(async () => {
      env.state.ended = 'document-replaced';
    });
    restoreWorkers.push(() => pending.mockRestore());
    expect(await env.clients.read.readParagraphs({ view: 'accepted' })).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
  });

  test('worker reads refuse if the invocation ends during the read and never flush', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    worker.read.mockImplementationOnce(async (request) => {
      await Promise.resolve();
      env.state.ended = 'document-replaced';
      return { ok: true, version: env.session.version(), view: request.view, paragraphs: [] };
    });
    expect(await env.clients.read.readParagraphs({ view: 'accepted' })).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
    expect(await env.clients.read.version()).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
    expect(await env.clients.read.readParagraphs({ view: 'accepted' })).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
    expect(worker.read).toHaveBeenCalledTimes(1);
    expect(worker.flush).not.toHaveBeenCalled();
    expect(worker.replica).not.toHaveBeenCalled();
  });

  for (const held of [true, false]) {
    test(`viewer text search reads ${held ? 'the worker with a held document' : 'the main copy already present'}`, async () => {
      const env = await setup();
      const worker = routeWorker(env);
      if (!held) worker.routing.mockReturnValue(null);
      env.state.viewer = !held;
      const release = mock(() => { throw new Error('unexpected viewer release'); });
      if (held) workerOpenReplica.holdWorkerOpenDocument(env.session, release);
      const awaitReplica = spyOn(workerOpenReplica, 'awaitWorkerOpenReplica');
      const ensureReplica = spyOn(workerOpenReplica, 'ensureWorkerOpenReplica');
      const mainFind = spyOn(env.session, 'findText');
      restoreWorkers.push(() => { awaitReplica.mockRestore(); ensureReplica.mockRestore(); mainFind.mockRestore(); });
      const request = { text: 'Tail', within: { kind: 'story', story: 'body' }, view: 'accepted' } as const;
      expect(await env.clients.read.findText(request)).toEqual({
        ok: true, version: env.session.version(), matches: held ? [] : expect.any(Array), truncated: false,
      });
      if (held) {
        expect(worker.findText).toHaveBeenCalledWith(request, expect.any(Function));
        expect(mainFind).not.toHaveBeenCalled();
        expect(worker.flush).not.toHaveBeenCalled();
        expect(worker.replica).not.toHaveBeenCalled();
      } else {
        expect(worker.findText).not.toHaveBeenCalled();
        expect(mainFind).toHaveBeenCalledWith(request);
        expect(worker.flush).toHaveBeenCalledTimes(1);
        expect(worker.replica).not.toHaveBeenCalled();
      }
      expect(awaitReplica).not.toHaveBeenCalled();
      expect(ensureReplica).not.toHaveBeenCalled();
      expect(release).not.toHaveBeenCalled();
    });
  }

  test('viewer text search preserves document replacement refusals during the worker read', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    env.state.viewer = true;
    workerOpenReplica.holdWorkerOpenDocument(env.session, () => { throw new Error('unexpected viewer release'); });
    worker.findText.mockImplementationOnce(async () => {
      env.pagedEditorRef.current = null;
      return { ok: true as const, version: env.session.version(), matches: [], truncated: false };
    });
    expect(await env.clients.read.findText({
      text: 'Tail', within: { kind: 'story', story: 'body' }, view: 'accepted',
    })).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
    expect(worker.flush).not.toHaveBeenCalled();
    expect(worker.replica).not.toHaveBeenCalled();
  });

  test('viewer text search refuses an authority main fallback without flushing or replica admission', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    env.state.viewer = true;
    workerOpenReplica.holdWorkerOpenDocument(env.session, () => { throw new Error('unexpected viewer release'); });
    worker.findText.mockImplementation((_request, main) => main());
    expect(await env.clients.read.findText({
      text: 'Tail', within: { kind: 'story', story: 'body' }, view: 'accepted',
    })).toMatchObject({ ok: false, failure: { code: 'input-failed' } });
    expect(worker.flush).not.toHaveBeenCalled();
    expect(worker.replica).not.toHaveBeenCalled();
  });

  test('text searches wait for hydration and return a refusal if the document is replaced', async () => {
    const env = await setup();
    const worker = routeWorker(env);
    let release!: () => void;
    const waiting = new Promise<void>((resolve) => (release = resolve));
    worker.replica.mockImplementation(() => waiting);
    const found = env.clients.read.findText({
      text: 'Tail',
      within: { kind: 'story', story: 'body' },
      view: 'accepted',
    });
    expect(worker.replica).toHaveBeenCalledWith(env.session);
    expect(worker.flush).not.toHaveBeenCalled();
    env.pagedEditorRef.current = null;
    release();
    expect(await found).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
    expect(worker.flush).not.toHaveBeenCalled();
  });

  test('without a worker authority reads and focus navigation still flush and never request a replica', async () => {
    const env = await setup();
    const replica = spyOn(workerOpenReplica, 'requestWorkerOpenReplica');
    restoreWorkers.push(() => replica.mockRestore());
    await env.clients.read.version();
    await env.clients.read.readParagraphs({ view: 'accepted' });
    await env.clients.read.findText({
      text: 'Tail',
      within: { kind: 'story', story: 'body' },
      view: 'accepted',
    });
    await env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: '00000002' },
      { expectVersion: env.session.version(), focus: true }
    );
    expect(env.events.filter((event) => event === 'flush')).toHaveLength(4);
    expect(env.events.slice(-2)).toEqual(['sync:false:*', 'focus']);
    expect(replica).not.toHaveBeenCalled();
  });

  test('reads after a flush and refuses once the invocation ends', async () => {
    const env = await setup();
    expect(await env.clients.read.version()).toEqual({ ok: true, version: env.session.version() });
    const read = await env.clients.read.readParagraphs({ view: 'accepted' });
    expect(read.ok && read.paragraphs.slice(0, 2).map((candidate) => candidate.text)).toEqual([
      'Alpha',
      'Tail',
    ]);
    env.state.ended = 'document-replaced';
    expect(
      await env.clients.read.findText({
        text: 'Tail',
        within: { kind: 'story', story: 'body' },
        view: 'accepted',
      })
    ).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
  });

  test('scrolls to a unique body paragraph without moving focus unless asked', async () => {
    const env = await setup();
    const version = env.session.version();
    const target = { story: 'body', paraId: '00000002' };
    const selection = JSON.stringify(env.session.selection());
    expect(
      await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })
    ).toEqual({
      ok: true,
    });
    const scrolled = env.events.filter((event) => !event.startsWith('flush'));
    expect(scrolled).toHaveLength(1);
    expect(scrolled[0]).toMatch(/^scroll:\d+$/);
    expect(JSON.stringify(env.session.selection())).toBe(selection);
    await env.clients.navigation.scrollToParagraph(target, { expectVersion: version, focus: true });
    expect(env.events.slice(-3)).toEqual([scrolled[0], 'sync:false:*', 'focus']);
    expect(env.session.selection()?.head).toMatchObject({ paraId: '00000002', offset: 0 });

    const failures = [
      [{ story: 'body', paraId: 'missing' }, version, 'missing-target'],
      [{ story: 'hf:rIdHeader', paraId: '0000E001' }, version, 'unsupported'],
      [target, 'stale', 'stale-version'],
    ] as const;
    for (const [where, expectVersion, code] of failures) {
      expect(
        await env.clients.navigation.scrollToParagraph(where, { expectVersion })
      ).toMatchObject({
        ok: false,
        failure: { code },
      });
    }
    for (const reveal of ['unsupported', 'layout-unavailable'] as const) {
      env.state.reveal = reveal;
      expect(
        await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })
      ).toMatchObject({ ok: false, failure: { code: reveal } });
    }
    env.state.reveal = 'scrolled';
    env.state.layoutReady = false;
    env.state.layoutFailed = true;
    const before = env.events.length;
    expect(
      await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })
    ).toMatchObject({
      ok: false,
      failure: { code: 'layout-unavailable' },
    });
    expect(env.events.slice(before).some((event) => event.startsWith('scroll'))).toBe(false);
  });

  test('scrolls when the expected layout arrives after 1.5 seconds', async () => {
    const env = await setup();
    const clock = navigationClock();
    const version = env.session.version();
    stampSourceVersion(env.queries, 'older');
    const scroll = env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: '00000002' },
      { expectVersion: version }
    );
    await env.waiting;
    clock.advance(1500);
    expect(clock.timers.size).toBe(1);
    expect(env.events.some((event) => event.startsWith('scroll'))).toBe(false);
    stampSourceVersion(env.queries, version);
    env.publishLayout();
    expect(await scroll).toEqual({ ok: true });
    expect(env.events.filter((event) => event.startsWith('scroll'))).toHaveLength(1);
    expect(clock.timers.size).toBe(0);
    expect(env.layoutListeners.size).toBe(0);
  });

  test('waits for the target page in a partial layout', async () => {
    const env = await setup();
    const clock = navigationClock();
    env.state.partial = true;
    env.state.anchorReady = false;
    const scroll = env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: '00000002' },
      { expectVersion: env.session.version(), focus: true }
    );
    await env.waiting;
    env.publishLayout();
    expect(env.events).toEqual(['flush']);
    expect(clock.timers.size).toBe(1);
    env.state.anchorReady = true;
    env.publishLayout();
    expect(await scroll).toEqual({ ok: true });
    expect(env.events.filter((event) => event.startsWith('scroll'))).toHaveLength(1);
    expect(env.events.slice(-2)).toEqual(['sync:false:*', 'focus']);
    expect(clock.timers.size).toBe(0);
  });

  test('a complete layout with no target position returns unsupported', async () => {
    const env = await setup();
    const clock = navigationClock();
    env.state.partial = true;
    env.state.anchorReady = false;
    const scroll = env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: '00000002' },
      { expectVersion: env.session.version() }
    );
    await env.waiting;
    env.state.partial = false;
    env.publishLayout();
    expect(await scroll).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
    expect(clock.timers.size).toBe(0);
  });

  test('a version change during the wait returns stale-version', async () => {
    const env = await setup();
    const clock = navigationClock();
    env.state.layoutReady = false;
    const scroll = env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: '00000002' },
      { expectVersion: env.session.version() }
    );
    await env.waiting;
    env.session.insertText({ story: 'body', paraId: '00000001', offset: 5 }, '!');
    expect(await scroll).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(clock.timers.size).toBe(0);
    expect(env.layoutListeners.size).toBe(0);
    expect(env.events).toEqual(['flush']);
  });

  test('closing the host aborts a wait through the plugin lifetime', async () => {
    const env = await setup();
    const host = createDocxPluginHost({
      ...env.access,
      geometry: () => null,
      translate: (key) => key,
    });
    let initialized!: (context: DocxPluginContext<null>) => void;
    const ready = new Promise<DocxPluginContext<null>>((resolve) => (initialized = resolve));
    host.setPlugins([
      defineDocxPlugin({
        id: 'acme.review',
        createState: () => null,
        initialize: (context) => initialized(context),
      }),
    ]);
    host.open(env.session);
    const context = await ready;
    const clock = navigationClock();
    try {
      env.state.layoutReady = false;
      const scroll = context.navigation.scrollToParagraph(
        { story: 'body', paraId: '00000002' },
        { expectVersion: env.session.version(), focus: true }
      );
      await env.waiting;
      expect(clock.timers.size).toBe(1);
      host.close('unmounted');
      expect(clock.timers.size).toBe(0);
      expect(env.layoutListeners.size).toBe(0);
      expect(await scroll).toMatchObject({ ok: false, failure: { code: 'plugin-unavailable' } });
      env.state.layoutReady = true;
      env.publishLayout();
      clock.advance(30_000);
      expect(env.events).toEqual(['flush']);
    } finally {
      host.close('unmounted');
    }
  });

  test('closing, unmounting and replacement cancel the wait and its timer', async () => {
    const endings = [
      ['closed', 'plugin-unavailable'],
      ['unmounted', 'plugin-unavailable'],
      ['document-replaced', 'document-replaced'],
    ] as const;
    for (const [reason, code] of endings) {
      const env = await setup();
      const clock = navigationClock();
      env.state.layoutReady = false;
      const scroll = env.clients.navigation.scrollToParagraph(
        { story: 'body', paraId: '00000002' },
        { expectVersion: env.session.version(), focus: true }
      );
      await env.waiting;
      expect(clock.timers.size).toBe(1);
      env.state.ended = code;
      if (reason === 'unmounted') env.pagedEditorRef.current = null;
      env.lifetimeController.abort();
      expect(clock.timers.size).toBe(0);
      expect(env.layoutListeners.size).toBe(0);
      expect(await scroll).toMatchObject({ ok: false, failure: { code } });
      env.state.layoutReady = true;
      env.publishLayout();
      clock.advance(30_000);
      expect(env.events).toEqual(['flush']);
      for (const restore of restoreClocks.splice(0)) restore();
    }
  });

  test('layout and query failures end the wait early', async () => {
    for (const failure of ['layout', 'queries'] as const) {
      const env = await setup();
      const clock = navigationClock();
      env.state.partial = true;
      env.state.anchorReady = false;
      const scroll = env.clients.navigation.scrollToParagraph(
        { story: 'body', paraId: '00000002' },
        { expectVersion: env.session.version() }
      );
      await env.waiting;
      if (failure === 'layout') env.state.layoutFailed = true;
      else env.state.queryState = 'error';
      env.publishLayout();
      expect(await scroll).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
      expect(clock.timers.size).toBe(0);
      expect(env.events).toEqual(['flush']);
      for (const restore of restoreClocks.splice(0)) restore();
    }
  });

  test('an unavailable layout reaches the 30-second safety cap', async () => {
    const env = await setup();
    const clock = navigationClock();
    env.state.layoutReady = false;
    const scroll = env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: '00000002' },
      { expectVersion: env.session.version() }
    );
    await env.waiting;
    clock.advance(29_999);
    expect(clock.timers.size).toBe(1);
    clock.advance(1);
    expect(await scroll).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
    expect(clock.timers.size).toBe(0);
    expect(env.layoutListeners.size).toBe(0);
    expect(env.events).toEqual(['flush']);
  });

  test('a newer call from another context supersedes the older wait for the same plugin', async () => {
    for (const layoutArrived of [false, true]) {
      const env = await setup();
      const clock = navigationClock();
      env.state.layoutReady = false;
      const version = env.session.version();
      const older = env.clients.navigation.scrollToParagraph(
        { story: 'body', paraId: '00000002' },
        { expectVersion: version, focus: true }
      );
      await env.waiting;
      env.state.layoutReady = true;
      if (layoutArrived) env.publishLayout();
      const newer = env.createClients().navigation.scrollToParagraph(
        { story: 'body', paraId: '00000001' },
        { expectVersion: version }
      );
      expect(await newer).toEqual({ ok: true });
      expect(await older).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
      env.publishLayout();
      const located = resolveParagraph(env.session, { story: 'body', paraId: '00000001' });
      expect(typeof located).not.toBe('string');
      expect(env.events.filter((event) => event.startsWith('scroll'))).toEqual([
        `scroll:${typeof located === 'string' ? -1 : located.position}`,
      ]);
      expect(env.events).not.toContain('focus');
      expect(clock.timers.size).toBe(0);
      expect(env.layoutListeners.size).toBe(0);
      for (const restore of restoreClocks.splice(0)) restore();
    }
  });

  test('a read never touches the session once the plugin ended, whenever that happens', async () => {
    for (let ticks = 0; ticks < 12; ticks += 1) {
      let env!: Awaited<ReturnType<typeof setup>>;
      const touched: Array<string | null> = [];
      env = await setup({
        flush: () => {
          void (async () => {
            for (let tick = 0; tick < ticks; tick += 1) await Promise.resolve();
            env.state.ended = 'document-replaced';
          })();
        },
      });
      const read = env.session.readParagraphs.bind(env.session);
      env.session.readParagraphs = (request) => {
        touched.push(env.state.ended);
        return read(request);
      };
      const result = await env.clients.read.readParagraphs({ view: 'accepted' });
      expect(touched.every((ended) => ended === null)).toBe(true);
      if (!result.ok) expect(result.failure.code).toBe('document-replaced');
    }
  });
});

test('a paragraph id present twice in its story is ambiguous', () => {
  const session = {
    storyIds: () => ['body'],
    hasStory: (story: string) => story === 'body',
    paragraphIdCount: () => 2,
  } as unknown as YrsSession;
  expect(resolveParagraph(session, { story: 'body', paraId: 'dup' })).toBe('ambiguous-target');
});
