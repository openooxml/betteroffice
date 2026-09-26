import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type DocxEditRequest, type YrsSession } from '@betteroffice/docx/yrs';
import type { PluginInvocation } from '../../../../shared/plugin-host/runtime';
import { UNAVAILABLE_DOCX_COMMANDS } from '../commands/createDocxCommandStore';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import { createPluginClients, resolveParagraph } from './createPluginClients';
import type { DocxPluginGrant, DocxPluginSnapshot } from './types';

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
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

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
      return state.reveal;
    },
    focus: () => events.push('focus'),
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor as PagedEditorRef | null };
  const state = {
    mode: 'editing' as EditorMode,
    grant: options.grant ?? ({ document: 'write', editBatches: true } as DocxPluginGrant),
    layoutReady: true,
    reveal: 'scrolled' as 'scrolled' | 'layout-unavailable' | 'unsupported',
    ended: null as 'plugin-unavailable' | 'document-replaced' | null,
  };
  const controller = new AbortController();
  const invocation: PluginInvocation<DocxPluginSnapshot> = {
    pluginId: 'acme.review',
    activation: {},
    snapshot: {} as DocxPluginSnapshot,
    signal: controller.signal,
    lifetimeSignal: controller.signal,
    state: () => null,
    setState: () => false,
    onCleanup: () => {},
    run: async () => {},
    commit: (write) => write(),
    refusal: () => state.ended ?? (controller.signal.aborted ? 'aborted' : null),
  };
  const clients = createPluginClients(
    invocation,
    {
      pagedEditorRef,
      writeMode: () => state.mode,
      commands: () => null,
      settledLayout: async () => state.layoutReady,
    },
    () => state.grant,
    UNAVAILABLE_DOCX_COMMANDS
  );
  return { session, events, pagedEditorRef, editor, state, controller, clients };
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
    const before = env.events.length;
    expect(
      await env.clients.navigation.scrollToParagraph(target, { expectVersion: version })
    ).toMatchObject({
      ok: false,
      failure: { code: 'layout-unavailable' },
    });
    expect(env.events.slice(before).some((event) => event.startsWith('scroll'))).toBe(false);
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
    paragraphs: () => [{ paraId: 'dup' }, { paraId: 'dup' }],
  } as unknown as YrsSession;
  expect(resolveParagraph(session, { story: 'body', paraId: 'dup' })).toBe('ambiguous-target');
});
