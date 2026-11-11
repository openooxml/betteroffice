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

const INNER = '<w:p w14:paraId="00000012"><w:r><w:rPr><w:i/></w:rPr><w:t>Inner</w:t></w:r></w:p>';
const BLOCKS = {
  table: '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
    `${INNER}</w:tc></w:tr></w:tbl>`,
  contentControl: `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${INNER}</w:sdtContent></w:sdt>`,
};

const innerItalic = (arm: CommandArm) => {
  for (const { story } of arm.engine.listStories().stories) {
    const segment = arm.engine.storySegments(story).find((entry) => entry.kind === 'text' && entry.text.includes('Inner'));
    if (segment?.kind === 'text') return Boolean(segment.attributes.italic);
  }
  throw new Error('no text segment holds Inner');
};

for (const [kind, block] of Object.entries(BLOCKS)) {
  const body = '<w:p w14:paraId="00000011"><w:r><w:rPr><w:i/></w:rPr><w:t>Intro</w:t></w:r></w:p>' + block +
    '<w:p w14:paraId="00000013"><w:r><w:t>After</w:t></w:r></w:p>';
  const open = async () => {
    const arm = await openCommandEngine([], { bodyXml: body, stylesXml: STYLES });
    arm.state.selection = selection('00000011', 0);
    return arm;
  };
  const extend = (arm: CommandArm, unit: 'character', times: number) => {
    for (let step = 0; step < times; step += 1) {
      execute(arm, command(arm, 'navigate', { unit, direction: 'forward', extend: true }));
    }
    return arm.state.selection!;
  };

  test(`clear over a selection ending at the end of the paragraph before a ${kind} leaves the text inside it`, async () => {
    const arm = await open();
    const before = extend(arm, 'character', 5);
    expect(before.focus).toMatchObject({ paraId: '00000011', offset: 5 });
    expect(execute(arm, command(arm, 'clearFormatting', { selection: before })).status).toBe('applied');
    expect(innerItalic(arm)).toBe(true);
    expect(Boolean(marksOf(arm, 'Intro').italic)).toBe(false);
  }, TIMEOUT);

  test(`clear at a caret after a ${kind} leaves the text inside it`, async () => {
    const arm = await open();
    expect(execute(arm, command(arm, 'clearFormatting', { selection: selection('00000013', 0) })).status).toBe('noop');
    expect(innerItalic(arm)).toBe(true);
  }, TIMEOUT);
}
