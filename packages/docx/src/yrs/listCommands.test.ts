import { afterEach, beforeAll, expect, test } from 'bun:test';
import type { YrsParagraphAttrs, YrsSession } from './index';
import { LIST_OFF_ATTRS } from './listCommands';
import { createdListKey } from './listDefinitionSave';
import {
  GAP, HEADING, LIST_LAYOUT, PLAIN, listCommandScenarios, listDocument, preloadListWasm, shownLeft, shownMarker,
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
    left: (text) => {
      arm.engine.layoutDocumentWithRegionsRetained(LIST_LAYOUT);
      return shownLeft(arm.engine.buildDisplayListFrame('{}', 0), text);
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

const REFUSED_WITHOUT_LIST_DEFINITIONS = 'unsupported-command';

test('worker: an engine that cannot ensure list definitions refuses a list toggle and changes no paragraph', async () => {
  const arm = await openCommandEngine(undefined, { bytes: listDocument('C'), clientId: ++clientId });
  const paragraph = arm.engine.paragraphs('body').find(entry => entry.text === PLAIN);
  expect(paragraph).toBeDefined();
  Reflect.deleteProperty(arm.engine, 'ensureListDefinition');
  const before = arm.engine.paragraphs('body');
  const outcome = (() => {
    try {
      return execute(arm, command(arm, 'toggleList', { selection: selection(paragraph!.paraId, 0), kind: 'bullet' }));
    } catch (error) {
      return { thrown: String(error) };
    }
  })();
  expect(outcome).toMatchObject({ status: 'refused', reason: REFUSED_WITHOUT_LIST_DEFINITIONS });
  expect(arm.engine.paragraphs('body')).toEqual(before);
  expect('ensureListDefinition' in arm.engine).toBe(false);
}, TIMEOUT);

test('worker: a paragraph-attribute command carrying a created list with a string numId in other is admitted, and at the top level refused', async () => {
  const arm = await openCommandEngine(undefined, { bytes: listDocument('C'), clientId: ++clientId });
  const paragraph = arm.engine.paragraphs('body').find(entry => entry.text === PLAIN);
  expect(paragraph).toBeDefined();
  const listDef = { key: '5:0', numId: '2' };
  expect(execute(arm, command(arm, 'setParagraphAttrs', { selection: selection(paragraph!.paraId, 0), attrs: { other: { listDef } } })))
    .toMatchObject({ status: 'applied' });
  expect(execute(arm, command(arm, 'setParagraphAttrs', {
    selection: selection(paragraph!.paraId, 0), attrs: { listDef } as YrsParagraphAttrs,
  }))).toMatchObject({ status: 'refused', reason: 'invalid-paragraph-attrs' });
  expect(execute(arm, command(arm, 'setParagraphAttrs', {
    selection: selection(paragraph!.paraId, 0), attrs: { other: { listDef: { key: '5:0', numId: 2 } } },
  }))).toMatchObject({ status: 'applied' });
}, TIMEOUT);

test('worker: bullets on, off and on again reuse the created list, and undo and redo bring that list back', async () => {
  const driver = await openWorker('C', { isolated: true });
  const keyOf = () => createdListKey(driver.paragraphs().find(paragraph => paragraph.text === PLAIN)?.properties ?? {});
  driver.toggle('bullet', PLAIN);
  const key = keyOf();
  driver.toggle('bullet', PLAIN);
  driver.toggle('bullet', PLAIN);
  expect(keyOf()).toBe(key);
  expect(driver.definitions()).toHaveLength(1);
  driver.undo();
  expect(keyOf()).toBeUndefined();
  driver.undo();
  expect(keyOf()).toBe(key);
  driver.undo();
  expect(keyOf()).toBeUndefined();
  expect(driver.definitions()).toEqual([]);
  driver.redo();
  driver.redo();
  driver.redo();
  expect(keyOf()).toBe(key);
  expect(driver.definitions()).toHaveLength(1);
  expect(key).toBeDefined();
}, TIMEOUT);

test('worker: a created list a tracked change still names is not handed to another paragraph', async () => {
  const arm = await openCommandEngine(undefined, { bytes: listDocument('C', { isolated: true }), clientId: ++clientId });
  const paragraph = (text: string) => arm.engine.paragraphs('body').find(entry => entry.text === text)!;
  const toggle = (text: string, suggesting?: { name: string; date: string }) => execute(arm, command(arm, 'toggleList', {
    selection: selection(paragraph(text).paraId, 0), kind: 'bullet', ...(suggesting ? { suggesting } : {}),
  }));
  expect(toggle(PLAIN)).toMatchObject({ status: 'applied' });
  const key = createdListKey(paragraph(PLAIN).properties);
  expect(toggle(PLAIN, { name: 'Writer', date: '2026-01-01T00:00:00Z' })).toMatchObject({ status: 'applied' });
  expect(toggle(GAP)).toMatchObject({ status: 'applied' });
  expect(createdListKey(paragraph(GAP).properties)).toBeDefined();
  expect(createdListKey(paragraph(GAP).properties)).not.toBe(key);
  expect(createdListKey(paragraph(PLAIN).properties)).toBeUndefined();
  expect(key).toBeDefined();
}, TIMEOUT);

test('worker: bullets reusing the created list of a deleted paragraph are undone first, and the next undo brings that paragraph back alone on its list', async () => {
  const target = 'Target paragraph QA';
  const arm = await openCommandEngine(undefined, {
    bytes: listDocument('C', { isolated: true, after: ['Middle paragraph QA', target] }), clientId: ++clientId,
  });
  const paragraphs = () => arm.engine.paragraphs('body');
  const paragraph = (text: string) => paragraphs().find(entry => entry.text === text)!;
  const naming = (key: string | undefined) => paragraphs().filter(entry => createdListKey(entry.properties) === key).map(entry => entry.text);
  const definitions = () => (JSON.parse(arm.native.created_list_definitions_json()) as Array<{ key: string }>).map(definition => definition.key);
  const toggle = (text: string) => expect(execute(arm, command(arm, 'toggleList', {
    selection: selection(paragraph(text).paraId, 0), kind: 'bullet',
  }))).toMatchObject({ status: 'applied' });
  const undo = () => expect(execute(arm, command(arm, 'undo', {}))).toMatchObject({ status: 'applied' });
  toggle(PLAIN);
  const key = createdListKey(paragraph(PLAIN).properties) ?? '';
  expect(execute(arm, command(arm, 'deleteBackward', { selection: {
    anchor: { story: 'body', paraId: paragraph(GAP).paraId, offset: GAP.length },
    focus: { story: 'body', paraId: paragraph(PLAIN).paraId, offset: PLAIN.length },
  } }))).toMatchObject({ status: 'applied' });
  const deleted = {
    texts: paragraphs().map(entry => entry.text),
    lists: paragraphs().filter(entry => entry.properties.listDef != null).length,
    definitions: definitions(),
  };
  toggle(target);
  const reused = { naming: naming(key), definitions: definitions() };
  undo();
  const undone = paragraph(target).properties;
  undo();
  expect(naming(key)).toEqual([PLAIN]);
  expect(createdListKey(paragraph(target).properties)).toBeUndefined();
  expect(paragraph(target).properties.numPr ?? null).toBeNull();
  expect(definitions()).toEqual([key]);
  expect(createdListKey(undone)).toBeUndefined();
  expect(undone.numPr ?? null).toBeNull();
  expect(reused).toEqual({ naming: [target], definitions: [key] });
  expect(deleted).toEqual({ texts: ['Opening paragraph QA', GAP, 'Middle paragraph QA', target], lists: 0, definitions: [key] });
  expect(key).not.toBe('');
}, TIMEOUT);

test('worker: bullets off on a heading that was given a style without numbering write no numId 0', async () => {
  const arm = await openCommandEngine(undefined, { bytes: listDocument('D'), clientId: ++clientId });
  const heading = () => arm.engine.paragraphs('body').find(entry => entry.text === HEADING)!;
  const at = () => selection(heading().paraId, 0);
  expect(execute(arm, command(arm, 'toggleList', { selection: at(), kind: 'bullet' }))).toMatchObject({ status: 'applied' });
  expect(execute(arm, command(arm, 'applyParagraphStyle', { selection: at(), styleId: 'Normal' }))).toMatchObject({ status: 'applied' });
  const styled = heading().properties;
  expect(execute(arm, command(arm, 'toggleList', { selection: at(), kind: 'bullet' }))).toMatchObject({ status: 'applied' });
  expect(heading().properties.numPr ?? null).toBeNull();
  expect(createdListKey(styled)).toBeDefined();
  expect(styled.pStyle).toBe('Normal');
}, TIMEOUT);

test('worker: a paragraph moved to another list by a paragraph-attribute write loses the list display attribute', async () => {
  const open = async (doc: 'G' | 'J') => {
    const arm = await openCommandEngine(undefined, { bytes: listDocument(doc), clientId: ++clientId });
    const plain = () => arm.engine.paragraphs('body').find(entry => entry.text === PLAIN)!;
    const at = () => selection(plain().paraId, 0);
    const toggle = (kind: 'bullet' | 'numbered') =>
      expect(execute(arm, command(arm, 'toggleList', { selection: at(), kind }))).toMatchObject({ status: 'applied' });
    const write = (numPr: { numId: number; ilvl: number }) =>
      expect(execute(arm, command(arm, 'setParagraphAttrs', { selection: at(), attrs: { numPr } }))).toMatchObject({ status: 'applied' });
    return { plain, toggle, write };
  };
  const moved = await open('J');
  moved.toggle('bullet');
  const joined = moved.plain().properties;
  moved.write({ numId: 5, ilvl: 0 });
  const replaced = moved.plain().properties;
  const nested = await open('G');
  nested.toggle('numbered');
  const level = nested.plain().properties;
  nested.write({ numId: 7, ilvl: 0 });
  const outdented = nested.plain().properties;
  nested.write({ numId: 2, ilvl: 0 });
  expect(replaced.numPr).toMatchObject({ numId: 5, ilvl: 0 });
  expect(replaced.listLevelIndents ?? null).toBeNull();
  expect(outdented.numPr).toMatchObject({ numId: 7, ilvl: 0 });
  expect(outdented.listLevelIndents).toBe(true);
  expect(nested.plain().properties.listLevelIndents ?? null).toBeNull();
  expect(joined.numPr).toMatchObject({ numId: 6, ilvl: 0 });
  expect(joined.listLevelIndents).toBe(true);
  expect(level.numPr).toMatchObject({ numId: 7, ilvl: 1 });
  expect(level.listLevelIndents).toBe(true);
}, TIMEOUT);

test('a session written before the style question still satisfies the session type', () => {
  const adapter = {} as Omit<YrsSession, 'paragraphStyleNumbers'>;
  const session: YrsSession = adapter;
  expect(session.paragraphStyleNumbers).toBeUndefined();
});

test('worker: an engine without the style question takes a style-numbered heading off by the flag stored at load', async () => {
  const arm = await openCommandEngine(undefined, { bytes: listDocument('D'), clientId: ++clientId });
  const heading = () => arm.engine.paragraphs('body').find(entry => entry.text === HEADING)!;
  Reflect.deleteProperty(arm.engine, 'paragraphStyleNumbers');
  expect(execute(arm, command(arm, 'toggleList', { selection: selection(heading().paraId, 0), kind: 'numbered' })))
    .toMatchObject({ status: 'applied' });
  expect(heading().properties.numPr).toMatchObject({ numId: 0 });
  expect('paragraphStyleNumbers' in arm.engine).toBe(false);
}, TIMEOUT);

test('the attributes that take a paragraph off its list are frozen and clear its marker formatting', () => {
  expect(Object.isFrozen(LIST_OFF_ATTRS)).toBe(true);
  expect(Object.isFrozen(LIST_OFF_ATTRS.other)).toBe(true);
  for (const key of ['listMarkerHidden', 'listMarkerBold', 'listMarkerItalic', 'listMarkerColor', 'listMarkerFontSize']) {
    expect(LIST_OFF_ATTRS).toHaveProperty(key, null);
  }
});
