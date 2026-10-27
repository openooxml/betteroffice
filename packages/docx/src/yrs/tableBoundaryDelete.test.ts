import { afterEach, expect, test } from 'bun:test';
import { decodeFrameDelta } from '../layout/render/frameDelta';
import {
  closeCommandEngines, command, execute, openCommandEngine, selection, TIMEOUT,
} from './residentCommandTestSupport';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const table = (prefix: string) => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${[1, 2].map(row => `<w:tr>${['A', 'B'].map(column => `<w:tc>${paragraph(`${prefix}-${column}${row}`)}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`;
const body = (between: boolean) => paragraph('Paragraph before the table.') + table('T1') +
  (between ? paragraph('') + table('T2') : '') + paragraph('Paragraph after the table.');
const layout = JSON.stringify({
  bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } }, renderEnv: {},
});

afterEach(closeCommandEngines);

for (const between of [false, true]) {
  test(`backspace ${between ? 'between tables' : 'after a table'} preserves the story and types in the last cell`, async () => {
    const arm = await openCommandEngine([], { bodyXml: body(between) });
    const paragraphs = arm.engine.paragraphs('body');
    const target = paragraphs[1]!;
    const before = arm.engine.storySegments('body');
    const cellStory = arm.engine.geometryReader.storyIds().find(story =>
      arm.engine.paragraphs(story).some(para => para.text === 'T1-B2'))!;
    const cell = arm.engine.paragraphs(cellStory)[0]!;
    const deleted = execute(arm, command(arm, 'deleteBackward', {
      selection: selection(target.paraId, 1),
    }));
    expect(arm.engine.storySegments('body')).toEqual(before);
    expect(deleted.status).toBe('noop');
    expect(arm.state.selection?.focus).toMatchObject({ story: cellStory, paraId: cell.paraId, offset: 5 });
    expect(execute(arm, command(arm, 'insertText', { text: 'Hello' })).status).toBe('applied');
    expect(arm.engine.paragraphs(cellStory)[0]!.text).toBe('T1-B2Hello');
    expect(() => arm.engine.layoutDocumentWithRegionsRetained(layout)).not.toThrow();
    expect(arm.engine.storySegments('body')).toEqual(before);
    expect(execute(arm, command(arm, 'undo', {})).status).toBe('applied');
    expect(arm.engine.paragraphs(cellStory)[0]!.text).toBe('T1-B2');
    expect(execute(arm, command(arm, 'redo', {})).status).toBe('applied');
    expect(arm.engine.paragraphs(cellStory)[0]!.text).toBe('T1-B2Hello');
    expect(arm.engine.storySegments('body')).toEqual(before);
  }, TIMEOUT);
}

test('forward delete before a table preserves the boundary and subsequent typing', async () => {
  const arm = await openCommandEngine([], { bodyXml: body(false) });
  const first = arm.engine.paragraphs('body')[0]!;
  const before = arm.engine.storySegments('body');
  const deleted = execute(arm, command(arm, 'deleteForward', {
    selection: selection(first.paraId, first.text.length),
  }));
  expect(deleted.status).toBe('noop');
  expect(arm.engine.storySegments('body')).toEqual(before);
  expect(execute(arm, command(arm, 'insertText', { text: 'Hello' })).status).toBe('applied');
  expect(arm.engine.paragraphs('body')[0]!.text).toBe(first.text + 'Hello');
  expect(() => arm.engine.layoutDocumentWithRegionsRetained(layout)).not.toThrow();
}, TIMEOUT);

test('a selected range spanning a table deletes the selection and undo restores the table', async () => {
  const arm = await openCommandEngine([], { bodyXml: body(false) });
  const [first, last] = arm.engine.paragraphs('body');
  const before = arm.engine.storySegments('body');
  const deleted = execute(arm, command(arm, 'deleteBackward', {
    selection: {
      anchor: { story: 'body', paraId: first!.paraId, offset: 3 },
      focus: { story: 'body', paraId: last!.paraId, offset: 4 },
    },
  }));
  expect(deleted.status).toBe('applied');
  expect(arm.engine.storySegments('body').filter(segment => segment.kind === 'embed' && segment.embedKind === 'table')).toHaveLength(0);
  expect(() => arm.engine.layoutDocumentWithRegionsRetained(layout)).not.toThrow();
  expect(execute(arm, command(arm, 'undo', {})).status).toBe('applied');
  expect(arm.engine.storySegments('body')).toEqual(before);
  expect(execute(arm, command(arm, 'redo', {})).status).toBe('applied');
  expect(arm.engine.storySegments('body').filter(segment => segment.kind === 'embed' && segment.embedKind === 'table')).toHaveLength(0);
}, TIMEOUT);

for (const backward of [true, false]) {
  test(`counted ${backward ? 'backward' : 'forward'} deletion stops at the table boundary`, async () => {
    const arm = await openCommandEngine([], { bodyXml: body(false) });
    const paragraphs = arm.engine.geometryReader.paragraphSpans('body');
    const target = paragraphs[backward ? 1 : 0]!;
    const removed = execute(arm, command(arm, backward ? 'deleteBackward' : 'deleteForward', {
      selection: selection(target.paraId, backward ? target.length : 0), count: 100,
    }));
    expect(removed.status).toBe('applied');
    expect(arm.engine.storySegments('body').filter(segment => segment.kind === 'embed' && segment.embedKind === 'table')).toHaveLength(1);
    expect(arm.engine.paragraphs('body')).toHaveLength(2);
    expect(() => arm.engine.layoutDocumentWithRegionsRetained(layout)).not.toThrow();
    expect(execute(arm, command(arm, 'insertText', { text: 'Hello' })).status).toBe('applied');
    expect(() => arm.engine.layoutDocumentWithRegionsRetained(layout)).not.toThrow();
  }, TIMEOUT);
}

test('resident counted backward deletion stops before a table embed', async () => {
  const arm = await openCommandEngine([], { bodyXml: body(false) });
  const target = arm.engine.geometryReader.paragraphSpans('body')[1]!;
  const at = { story: 'body', paraId: target.paraId, offset: target.length };
  arm.engine.setSelection(at, at);
  const frame = decodeFrameDelta(arm.engine.buildDisplayListFrame('{}', 0));
  arm.resident.applyDelete('backward', frame.frameEpoch, 100);
  expect(arm.engine.storySegments('body').filter(segment => segment.kind === 'embed' && segment.embedKind === 'table')).toHaveLength(1);
  expect(arm.engine.paragraphs('body')).toHaveLength(2);
  expect(arm.resident.residentDeletedUnits()).toBe(target.length - 1);
  expect(() => arm.engine.layoutDocumentWithRegionsRetained(layout)).not.toThrow();
}, TIMEOUT);

test('resident backspace at a table refuses before mutating the story', async () => {
  const arm = await openCommandEngine([], { bodyXml: body(false) });
  const target = arm.engine.geometryReader.paragraphSpans('body')[1]!;
  const at = { story: 'body', paraId: target.paraId, offset: 1 };
  arm.engine.setSelection(at, at);
  const frame = decodeFrameDelta(arm.engine.buildDisplayListFrame('{}', 0));
  const before = arm.engine.storySegments('body');
  expect(() => arm.resident.applyDelete('backward', frame.frameEpoch, 1)).toThrow('resident input state is not ready');
  expect(arm.engine.storySegments('body')).toEqual(before);
}, TIMEOUT);
