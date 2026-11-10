import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '@betteroffice/docx/docx/rezip/parts';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsCellLoc, type YrsSession } from '@betteroffice/docx/yrs';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { useRustDisplayList, type ResidentFrameApplyResult } from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;

const path = (relative: string) => fileURLToPath(new URL(relative, import.meta.url));
const WASM = path('../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = path('../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});
const sessions: YrsSession[] = [];
let startWorker: () => InProcessResidentWorker;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  startWorker = await residentWorkerFactory();
}, 30_000);

afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
  for (const session of sessions.splice(0)) session.destroy();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;
const body = paragraph('Before') +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid><w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>` +
  paragraph('After');

function docx(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const at = (row: number, column: number): YrsCellLoc => ({ story: 'body', tableIndex: 0, row, column });

function sync(a: YrsSession, b: YrsSession): void {
  a.applyUpdate(b.encodeStateAsUpdate(a.encodeStateVector()));
  b.applyUpdate(a.encodeStateAsUpdate(b.encodeStateVector()));
}

function danglingCells(session: YrsSession): string[] {
  const stories = session.storyIds();
  const missing: string[] = [];
  for (const story of stories) {
    for (const segment of session.storySegments(story)) {
      if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
      for (const row of (segment.payload.rows as { cells?: { story?: string }[] }[] | undefined) ?? []) {
        for (const named of row.cells ?? []) if (named.story && !stories.includes(named.story)) missing.push(named.story);
      }
    }
  }
  return missing;
}

const stateBytes = (session: YrsSession) =>
  `${Array.from(session.encodeStateVector()).join(',')}|${Array.from(session.encodeStateAsUpdate()).join(',')}`;

/** A laid-out host replica whose live table names a missing cell. */
async function danglingHost(base: number, order: 'synced' | 'concurrent'): Promise<{ host: YrsSession; story: string }> {
  const a = await createYrsSession({ clientId: base + 1 });
  sessions.push(a);
  a.openDocx(docx(), true);
  a.beginUndoCapture();
  const b = await createYrsSession({ clientId: base + 2 });
  sessions.push(b);
  b.openDocx(docx(), false);
  b.loadState(a.encodeState());
  b.beginUndoCapture();
  a.splitCell(at(0, 1), 1, 2);
  sync(a, b);
  b.insertRow(at(0, 0), 'below');
  if (order === 'synced') sync(a, b);
  expect(a.undo()).toBe(true);
  sync(a, b);
  const host = await createYrsSession({ clientId: base + 3 });
  sessions.push(host);
  host.openDocx(docx(), false);
  host.loadState(b.encodeState());
  host.registerFont(new Uint8Array(readFileSync(FONT)));
  const [story] = danglingCells(host);
  expect(story).toBeDefined();
  return { host, story: story! };
}

for (const [index, order] of (['synced', 'concurrent'] as const).entries()) {
  test(`${order}: typing in a real missing cell through the hook and a real worker is declined without an error, a fallback or a new worker`, async () => {
    const { host, story } = await danglingHost(9300 + index * 10, order);
    const inputs = JSON.parse(host.layoutDocumentWithRegionsJson(LAYOUT));
    let created = 0;
    globalThis.Worker = class {
      constructor() {
        created += 1;
        return startWorker();
      }
    } as unknown as typeof Worker;
    const errors = spyOn(console, 'error').mockImplementation(() => {});
    try {
      const overrides = { getInputs: () => inputs };
      const { result, unmount } = renderHook(
        ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, host),
        { initialProps: { layout: inputs.layout as Layout } }
      );
      await waitFor(() => {
        if (result.current.error) throw result.current.error;
        expect(result.current.frame).not.toBeNull();
        expect(result.current.workerSurfacesActive).toBe(true);
      }, { timeout: 20_000 });
      const caret = { story, paraId: host.paragraphSpans(story)[0]!.paraId, offset: 0 };
      host.setSelection(caret, caret);
      const before = stateBytes(host);
      let typed: ResidentFrameApplyResult | null | undefined;
      await act(async () => {
        typed = await result.current.applyInput('x');
      });
      await act(async () => {
        await result.current.applyDelete('backward');
      });
      expect(typed).toBeNull();
      expect(stateBytes(host) === before).toBe(true);
      expect(result.current.error).toBeNull();
      expect(result.current.workerSurfacesActive).toBe(true);
      expect(created).toBe(1);
      expect(danglingCells(host)).toContain(story);

      const after = host.paragraphSpans('body').at(-1)!;
      host.setSelection({ story: 'body', paraId: after.paraId, offset: after.length });
      let elsewhere: ResidentFrameApplyResult | null | undefined;
      await act(async () => {
        elsewhere = await result.current.applyInput(' typed');
      });
      expect(elsewhere?.frameEpoch ?? null).not.toBeNull();
      expect(host.paragraphs('body').some((entry) => entry.text === 'After typed')).toBe(true);
      expect(created).toBe(1);
      expect(errors.mock.calls.filter(([message]) => /\[(CanvasRenderer|YrsInput)\]/.test(String(message)))).toEqual([]);
      unmount();
    } finally {
      errors.mockRestore();
    }
  }, 60_000);
}
