import { afterEach, expect, test } from 'bun:test';
import {
  TIMEOUT, closeCommandEngines, command, execute, openCommandEngine, selection, type CommandArm,
} from './residentCommandTestSupport';

afterEach(closeCommandEngines);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STYLES = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>`;
const BODY = '<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>Title</w:t></w:r></w:p>' +
  '<w:p w14:paraId="00000002"><w:r><w:t>Body</w:t></w:r></w:p>';

const marksOf = (arm: CommandArm, marker: string) => {
  const segment = arm.engine.storySegments('body').find((entry) => entry.kind === 'text' && entry.text.includes(marker));
  if (segment?.kind !== 'text') throw new Error(`no text segment holds ${marker}`);
  return segment.attributes;
};

test('the worker clearFormatting command keeps a heading bold and drops its direct italic', async () => {
  const arm = await openCommandEngine([], { bodyXml: BODY, stylesXml: STYLES });
  expect(execute(arm, command(arm, 'clearFormatting', { selection: selection('00000001', 0, 5) })).status).toBe('applied');
  expect(Boolean(marksOf(arm, 'Title').bold)).toBe(true);
  expect(Boolean(marksOf(arm, 'Title').italic)).toBe(false);
  expect(execute(arm, command(arm, 'insertText', { text: '!', selection: selection('00000002', 4) })).status).toBe('applied');
}, TIMEOUT);

test('text typed after clearing at a collapsed caret in a heading is heading-bold', async () => {
  const arm = await openCommandEngine([], { bodyXml: BODY, stylesXml: STYLES });
  expect(execute(arm, command(arm, 'clearFormatting', { selection: selection('00000001', 5) })).status).toBe('noop');
  expect(execute(arm, command(arm, 'insertText', { text: 'Typed', selection: selection('00000001', 5) })).status).toBe('applied');
  const typed = marksOf(arm, 'Typed');
  expect(Boolean(typed.bold)).toBe(true);
  expect(Boolean(typed.italic)).toBe(false);
  expect(execute(arm, command(arm, 'insertText', { text: '!', selection: selection('00000002', 4) })).status).toBe('applied');
}, TIMEOUT);
