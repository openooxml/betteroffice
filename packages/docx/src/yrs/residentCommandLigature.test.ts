import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, command, execute, expectCaret, openCommandEngine, selection, texts, TIMEOUT, type CommandArm } from './residentCommandTestSupport';

const asking = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w14:ligatures w14:val="standard"/></w:rPr><w:t>office</w:t></w:r></w:p>';

afterEach(closeCommandEngines);

const step = (arm: CommandArm, direction: 'forward' | 'backward') =>
  execute(arm, command(arm, 'navigate', { unit: 'character', direction, extend: false })).selection?.focus.offset;

test('arrow keys step through each letter of an asking "ffi" once, and Backspace after the second letter deletes one letter', async () => {
  const arm = await openCommandEngine([], { bodyXml: asking });
  const paraId = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, command(arm, 'select', { selection: selection(paraId, 1) })).status).toBe('noop');
  expect([step(arm, 'forward'), step(arm, 'forward'), step(arm, 'forward')]).toEqual([2, 3, 4]);
  expect([step(arm, 'backward'), step(arm, 'backward'), step(arm, 'backward')]).toEqual([3, 2, 1]);
  expect(execute(arm, command(arm, 'select', { selection: selection(paraId, 3) })).status).toBe('noop');
  const deleted = execute(arm, command(arm, 'deleteBackward', {}));
  expect(deleted.status).toBe('applied');
  expectCaret(deleted, { story: 'body', paraId, offset: 2 });
  expect(texts(arm)).toEqual(['ofice']);
}, TIMEOUT);
