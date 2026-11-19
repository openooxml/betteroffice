import { afterEach, beforeAll, expect, test } from 'bun:test';
import type { YrsParagraphAttrs } from './index';
import { LIST_OFF_ATTRS } from './listCommands';
import { createdListKey } from './listDefinitionSave';
import {
  LIST_LAYOUT, PLAIN, listCommandScenarios, listDocument, preloadListWasm, shownMarker,
  type ListCommandDriver, type ListDocument, type ListDocumentOptions,
} from './listCommandsTestSupport';
import { closeCommandEngines, command, execute, openCommandEngine, selection, TIMEOUT } from './residentCommandTestSupport';

beforeAll(preloadListWasm, 60_000);
afterEach(closeCommandEngines);

let clientId = 7400;

async function openWorker(doc: ListDocument, options?: ListDocumentOptions): Promise<ListCommandDriver> {
  const bytes = listDocument(doc, options);
  const arm = await openCommandEngine(undefined, { bytes, clientId: ++clientId });
  const paraId = (text: string) => arm.engine.paragraphs('body').find(paragraph => paragraph.text === text)?.paraId ?? '';
  const select = (from: string, to = from) => {
    execute(arm, command(arm, 'select', { selection: {
      anchor: { story: 'body', paraId: paraId(from), offset: 0 },
      focus: { story: 'body', paraId: paraId(to), offset: to === from ? 0 : to.length },
    } }));
  };
  const run = (intent: Parameters<typeof execute>[1]) => expect(execute(arm, intent)).toMatchObject({ status: 'applied' });
  return {
    toggle: (kind, from, to) => {
      select(from, to);
      run(command(arm, 'toggleList', { kind }));
    },
    indent: (direction, at) => {
      select(at);
      run(command(arm, 'changeIndent', { direction }));
    },
    undo: () => run(command(arm, 'undo', {})),
    redo: () => run(command(arm, 'redo', {})),
    canUndo: () => arm.engine.canUndo(),
    paragraphs: () => arm.engine.paragraphs('body'),
    definitions: () => JSON.parse(arm.native.created_list_definitions_json()) as Array<{ key: string; kind: string }>,
    marker: (text) => {
      arm.engine.layoutDocumentWithRegionsRetained(LIST_LAYOUT);
      return shownMarker(arm.engine.buildDisplayListFrame('{}', 0), text);
    },
    save: async () => new Uint8Array(await arm.engine.save(bytes.slice(), arm.hostJson ?? '', undefined, [], { full: false })),
  };
}

listCommandScenarios('worker', openWorker, TIMEOUT);

test('worker: a paragraph-attribute command carrying the created list a paragraph is on is admitted', async () => {
  const arm = await openCommandEngine(undefined, { bytes: listDocument('C'), clientId: ++clientId });
  const paragraph = arm.engine.paragraphs('body').find(entry => entry.text === PLAIN);
  expect(paragraph).toBeDefined();
  const listDef = { key: `${clientId}:9`, numId: 1 };
  const attrs = { numPr: { numId: 1, ilvl: 0 }, listDef } as YrsParagraphAttrs;
  expect(execute(arm, command(arm, 'setParagraphAttrs', { selection: selection(paragraph!.paraId, 0), attrs })))
    .toMatchObject({ status: 'applied' });
  const after = arm.engine.paragraphs('body').find(entry => entry.text === PLAIN);
  expect(after?.properties.listDef).toEqual(listDef);
  expect(execute(arm, command(arm, 'setParagraphAttrs', {
    selection: selection(paragraph!.paraId, 0), attrs: { numPr: null, listDef: null } as YrsParagraphAttrs,
  }))).toMatchObject({ status: 'applied' });
  expect(arm.engine.paragraphs('body').find(entry => entry.text === PLAIN)?.properties.listDef ?? null).toBeNull();
}, TIMEOUT);

test('worker: a list paragraph the toolbar made keeps its created list when it takes a paragraph style without numbering', async () => {
  const arm = await openCommandEngine(undefined, { bytes: listDocument('C'), clientId: ++clientId });
  const plain = () => arm.engine.paragraphs('body').find(entry => entry.text === PLAIN)!;
  const at = selection(plain().paraId, 0);
  expect(execute(arm, command(arm, 'toggleList', { selection: at, kind: 'bullet' }))).toMatchObject({ status: 'applied' });
  const key = createdListKey(plain().properties);
  expect(key).toBeDefined();
  expect(execute(arm, command(arm, 'applyParagraphStyle', { selection: at, styleId: 'Heading1' }))).toMatchObject({ status: 'applied' });
  const styled = plain().properties;
  expect(styled.pStyle).toBe('Heading1');
  expect(styled.numPr).toMatchObject({ numId: 1, ilvl: 0 });
  expect(createdListKey(styled)).toBe(key);
}, TIMEOUT);

test('the attributes that take a paragraph off its list are frozen and clear its marker formatting', () => {
  expect(Object.isFrozen(LIST_OFF_ATTRS)).toBe(true);
  expect(Object.isFrozen(LIST_OFF_ATTRS.other)).toBe(true);
  for (const key of ['listMarkerHidden', 'listMarkerBold', 'listMarkerItalic', 'listMarkerColor', 'listMarkerFontSize']) {
    expect(LIST_OFF_ATTRS).toHaveProperty(key, null);
  }
});
