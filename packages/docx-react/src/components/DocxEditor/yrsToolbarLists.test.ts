import { afterEach, beforeAll, expect } from 'bun:test';
import {
  createYrsInputPositionMap,
  createYrsSession,
  saveYrsDocx,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  LIST_LAYOUT,
  listCommandScenarios,
  listDocument,
  listFont,
  preloadListWasm,
  shownLeft,
  shownMarker,
  type ListCommandDriver,
  type ListDocument,
  type ListDocumentOptions,
} from '../../../../docx/src/yrs/listCommandsTestSupport';
import { applyYrsToolbarFormatting } from './yrsToolbar';

const sessions: YrsSession[] = [];
let clientId = 7600;

beforeAll(preloadListWasm, 60_000);
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

async function openSession(doc: ListDocument, options?: ListDocumentOptions): Promise<ListCommandDriver> {
  const session = await createYrsSession({ clientId: ++clientId });
  sessions.push(session);
  session.openDocx(listDocument(doc, options), true);
  session.beginUndoCapture();
  session.registerFont(listFont());
  const paraId = (text: string) =>
    session.paragraphs('body').find((paragraph) => paragraph.text === text)?.paraId ?? '';
  const map = () =>
    createYrsInputPositionMap(
      'body',
      session.paragraphs('body').map((paragraph) => ({
        paraId: paragraph.paraId,
        length: paragraph.text.length,
      }))
    );
  const select = (from: string, to = from) =>
    session.setSelection(
      { story: 'body', paraId: paraId(from), offset: 0 },
      { story: 'body', paraId: paraId(to), offset: to === from ? 0 : to.length }
    );
  return {
    toggle: (kind, from, to) => {
      select(from, to);
      expect(applyYrsToolbarFormatting(session, map(), kind === 'bullet' ? 'bulletList' : 'numberedList')).toBe(true);
    },
    indent: (direction, at) => {
      select(at);
      expect(applyYrsToolbarFormatting(session, map(), direction === 'increase' ? 'indent' : 'outdent')).toBe(true);
    },
    undo: () => expect(session.undo()).toBe(true),
    redo: () => expect(session.redo()).toBe(true),
    canUndo: () => session.canUndo(),
    paragraphs: () => session.paragraphs('body'),
    definitions: () => session.createdListDefinitions(),
    marker: (text) => {
      session.layoutDocumentWithRegionsRetainedJson(LIST_LAYOUT);
      return shownMarker(session.buildDisplayListFrame('{}', 0), text);
    },
    left: (text) => {
      session.layoutDocumentWithRegionsRetainedJson(LIST_LAYOUT);
      return shownLeft(session.buildDisplayListFrame('{}', 0), text);
    },
    save: async () => (await saveYrsDocx(session, { updateModifiedDate: false })).bytes,
  };
}

listCommandScenarios('session', openSession, 30_000);
