import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import {
  createDisplayListQueries,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { preloadLayoutWasm } from '@betteroffice/docx/wasm/layout';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { createPluginGeometry, pluginLayout } from '../../../plugins/geometry';
import { markPresented, stampSourceVersion } from './layoutProvenance';
import {
  positionAtClientPoint,
  resolvePointPosition,
  type PointPositionEditor,
} from './pointPosition';
import { createYrsPositionProjection, projectYrsDisplayPosition } from './yrsPositionProjection';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RAW = 'Keep gone words here';
const DELETED = { start: 5, end: 10 };
const sessions: YrsSession[] = [];
const disposables: DisplayListQueries[] = [];

function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">Keep </w:t></w:r><w:del w:id="1" w:author="Ann" w:date="2026-09-01T00:00:00Z"><w:r><w:delText xml:space="preserve">gone </w:delText></w:r></w:del><w:r><w:t>words here</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Tail</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
    )
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function wasm(path: string): Uint8Array {
  const generated = resolve(import.meta.dir, '../../../../../docx/src/wasm/generated');
  return new Uint8Array(readFileSync(resolve(generated, path)));
}

beforeAll(async () => {
  await preloadEditWasm(wasm('edit/docx_edit_bg.wasm'));
  await preloadLayoutWasm(wasm('layout/docx_layout_bg.wasm'));
});
afterEach(() => {
  for (const queries of disposables.splice(0)) queries.dispose();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** The paged editor's mapping: region-aware projection, then the story's input position map. */
function editorFor(session: YrsSession, hasPendingInput = () => false): PointPositionEditor {
  return {
    getYrsSession: () => session,
    hasPendingInput,
    displayPositionToYrsLoc(hit) {
      const target = projectYrsDisplayPosition(hit, (root) =>
        createYrsPositionProjection(session, root)
      );
      if (!target) return null;
      const map = createYrsInputPositionMap(target.story, session.paragraphSpans(target.story));
      return displayPositionToYrsLoc(map, target.displayPosition);
    },
  };
}

/**
 * A frame of the session's current version, the paragraph's markup text on one line, which its
 * host shows unless `presented` is false.
 */
async function paint(session: YrsSession, { stamp = true, presented = true } = {}) {
  const start = createYrsPositionProjection(session, 'body')!.positionForLoc({
    story: 'body',
    paraId: '00000001',
    offset: 0,
  })!;
  const queries = createDisplayListQueries({
    pages: [
      {
        pageIndex: 0,
        width: 800,
        height: 1000,
        contentBounds: { x: 80, y: 80, width: 640, height: 760 },
        primitives: [
          {
            kind: 'text',
            text: RAW,
            x: 100,
            baselineY: 150,
            width: 200,
            font: '400 16px Calibri',
            color: '#000000',
            docStart: start,
            docEnd: start + RAW.length,
            blockId: start,
            lineIndex: 0,
          },
        ],
      },
    ],
  });
  await queries.whenReady();
  disposables.push(queries);
  if (stamp) stampSourceVersion(queries, session.version());
  const host = document.createElement('div');
  host.className = 'canvas-pages';
  host.getBoundingClientRect = () => new DOMRect(40, 60, 848, 1048);
  const page = document.createElement('div');
  page.className = 'canvas-page';
  page.dataset.pageIndex = '0';
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  canvas.getBoundingClientRect = () => new DOMRect(64, 84, 800, 1000);
  page.getBoundingClientRect = canvas.getBoundingClientRect;
  page.append(canvas);
  host.append(page);
  if (presented) markPresented(host, queries.displayList);
  const point = (x: number, y: number) => ({ clientX: 64 + x, clientY: 84 + y });
  return { queries, host, start, point };
}

function texts(session: YrsSession, view: 'accepted' | 'original'): string {
  const read = session.readParagraphs({ view });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs[0]!.text;
}

async function open() {
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(fixture(), true);
  return session;
}

/** Plugin geometry over `queries` as the plugin host builds it. */
function pluginGeometry(
  editor: PointPositionEditor,
  session: YrsSession,
  host: HTMLElement,
  queries: DisplayListQueries,
  current = () => true
) {
  const dom = createRenderedDomContext(host, 1, {
    displayListQueries: queries,
    projector: createCanvasHostProjector(host, queries, 1),
  });
  const layout = pluginLayout(queries, session.version(), 1)!;
  const geometry = createPluginGeometry(
    layout,
    dom,
    document.createElement('div'),
    current,
    (hit) => resolvePointPosition(editor, hit, host, queries)
  );
  return { layout, geometry };
}

describe('point positions as edit targets', () => {
  test('a dropped point after a pending deletion inserts exactly there', async () => {
    const session = await open();
    expect(texts(session, 'accepted')).toBe('Keep words here');
    const { queries, host, start, point } = await paint(session);
    const { clientX, clientY } = point(225, 145);
    const position = positionAtClientPoint(editorFor(session), host, queries, 1, clientX, clientY);
    expect(position).not.toBeNull();
    const raw = position!.position - start;
    expect(raw).toBeGreaterThan(DELETED.end);
    expect(raw).toBeLessThan(RAW.indexOf(' here'));
    const offset = raw - (DELETED.end - DELETED.start);
    expect(position).toMatchObject({
      region: 'body',
      pageIndex: 0,
      version: session.version(),
      target: {
        kind: 'range',
        story: 'body',
        start: { paraId: '00000001', offset },
        end: { paraId: '00000001', offset },
        view: 'accepted',
      },
    });
    expect(session.selection()).toBeNull();

    const result = session.applyEdits({
      expectVersion: position!.version,
      steps: [{ op: 'insertText', target: position!.target, at: 'start', text: 'X' }],
    });
    expect(result).toMatchObject({ ok: true, applied: true });
    expect(texts(session, 'accepted')).toBe(
      `${'Keep words here'.slice(0, offset)}X${'Keep words here'.slice(offset)}`
    );
    expect(texts(session, 'original')).toBe(`${RAW.slice(0, raw)}X${RAW.slice(raw)}`);
  });

  test('a layout behind the document refuses instead of answering a shifted position', async () => {
    const session = await open();
    const editor = editorFor(session);
    const { queries, host, point } = await paint(session);
    const { clientX, clientY } = point(225, 145);
    const before = positionAtClientPoint(editor, host, queries, 1, clientX, clientY)!;
    expect(before).not.toBeNull();
    expect(positionAtClientPoint(editor, host, queries, 1, 20, 20)).toBeNull();

    session.insertText({ story: 'body', paraId: '00000001', offset: 0 }, 'Now ');
    expect(positionAtClientPoint(editor, host, queries, 1, clientX, clientY)).toBeNull();
    const dom = createRenderedDomContext(host, 1, {
      displayListQueries: queries,
      projector: createCanvasHostProjector(host, queries, 1),
    });
    const hit = dom.getPositionAtPoint(clientX, clientY);
    expect(hit).not.toBeNull();
    expect(resolvePointPosition(editor, hit, host, queries)).toBeNull();
    const refused = session.applyEdits({
      expectVersion: before.version,
      steps: [{ op: 'insertText', target: before.target, at: 'start', text: 'X' }],
    });
    expect(refused).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(texts(session, 'original')).toBe(`Now ${RAW}`);

    const unstamped = await paint(session, { stamp: false });
    expect(
      positionAtClientPoint(editor, unstamped.host, unstamped.queries, 1, clientX, clientY)
    ).toBeNull();
  });

  test('plugin geometry resolves through the same hit test and ends with its layout', async () => {
    const session = await open();
    const editor = editorFor(session);
    const { queries, host, point } = await paint(session);
    const { clientX, clientY } = point(225, 145);
    let current = true;
    const { layout, geometry } = pluginGeometry(editor, session, host, queries, () => current);
    const expected = positionAtClientPoint(editor, host, queries, 1, clientX, clientY)!;
    expect(geometry.getPositionAtPoint(clientX, clientY)).toEqual({
      ...expected,
      layoutId: layout.id,
    });
    expect(geometry.getPositionAtPoint(20, 20)).toBeNull();
    current = false;
    expect(geometry.getPositionAtPoint(clientX, clientY)).toBeNull();
    current = true;
    session.insertText({ story: 'body', paraId: '00000002', offset: 0 }, 'Z');
    expect(geometry.getPositionAtPoint(clientX, clientY)).toBeNull();
  });

  test('a layout the pages have not painted yet refuses until they show it', async () => {
    const session = await open();
    const editor = editorFor(session);
    const shown = await paint(session);
    session.insertText({ story: 'body', paraId: '00000002', offset: 0 }, 'Z');
    const next = await paint(session, { presented: false });
    const { clientX, clientY } = next.point(225, 145);
    markPresented(next.host, shown.queries.displayList);
    const { layout, geometry } = pluginGeometry(editor, session, next.host, next.queries);
    expect(layout.version).toBe(session.version());
    expect(positionAtClientPoint(editor, next.host, next.queries, 1, clientX, clientY)).toBeNull();
    expect(geometry.getPositionAtPoint(clientX, clientY)).toBeNull();

    markPresented(next.host, next.queries.displayList);
    const position = positionAtClientPoint(editor, next.host, next.queries, 1, clientX, clientY);
    expect(position).toMatchObject({ version: session.version() });
    expect(geometry.getPositionAtPoint(clientX, clientY)).toEqual({
      ...position!,
      layoutId: layout.id,
    });
  });

  test('input on its way to the session refuses until it lands', async () => {
    const session = await open();
    let pending = true;
    const editor = editorFor(session, () => pending);
    const { queries, host, point } = await paint(session);
    const { clientX, clientY } = point(225, 145);
    const { geometry } = pluginGeometry(editor, session, host, queries);
    expect(positionAtClientPoint(editor, host, queries, 1, clientX, clientY)).toBeNull();
    expect(geometry.getPositionAtPoint(clientX, clientY)).toBeNull();

    pending = false;
    expect(positionAtClientPoint(editor, host, queries, 1, clientX, clientY)).toMatchObject({
      version: session.version(),
    });
    expect(geometry.getPositionAtPoint(clientX, clientY)).not.toBeNull();
  });
});
