import { describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import * as yrs from './index';
import type { DocxEditStep } from './edits';
import { editSaveReopen, holdsProperty, savedRunProperties, textRunFixture } from './__fixtures__/textRuns';
import { setupStructuralSaves } from './structuralSaveTestSupport';

setupStructuralSaves();

const paragraph = (paraId: string) => ({ kind: 'paragraph', story: 'body', paraId }) as const;
const format = (paraId: string, rest: Record<string, unknown>) =>
  ({ op: 'formatText', target: paragraph(paraId), ...rest }) as DocxEditStep;
const textOf = (segments: ReturnType<yrs.YrsSession['storySegments']>, marker: string) => {
  const segment = segments.find((entry) => entry.kind === 'text' && entry.text.includes(marker));
  return segment?.kind === 'text' ? segment.attributes : {};
};

describe('formatText clear', () => {
  test('toolbar clear formatting on a heading and a character style run: saved rPr and live state', async () => {
    const session = await yrs.createYrsSession({ clientId: 71 });
    const reopened = await yrs.createYrsSession({ clientId: 72 });
    try {
      session.seedFromDocx(await textRunFixture());
      for (const [paraId, length] of [['00000003', 5], ['00000004', 10]] as const) {
        session.clearFormatting({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: length } });
      }
      const live = session.storySegments('body');
      const saved = await yrs.saveYrsDocx(session, { updateModifiedDate: false });
      reopened.seedFromDocx(saved.bytes);
      const xml = await (await JSZip.loadAsync(saved.bytes)).file('word/document.xml')!.async('string');
      for (const marker of ['Title', 'Strong run']) {
        console.log(`toolbar clear ${marker}: live bold=${Boolean(textOf(live, marker).bold)} saved rPr=${savedRunProperties(xml, marker)} reopened bold=${Boolean(textOf(reopened.storySegments('body'), marker).bold)}`);
      }
      expect(xml).toContain('Title');
    } finally {
      session.destroy();
      reopened.destroy();
    }
  });

  test('a cleared heading run inherits bold: no w:b in its saved rPr, bold after reopen and a second save', async () => {
    const { xml, resavedXml, reopened } = await editSaveReopen(yrs, [format('00000003', { clear: ['bold'] })]);
    for (const document of [xml, resavedXml]) {
      expect(savedRunProperties(document, 'Title') ?? '').not.toContain('<w:b');
    }
    expect(resavedXml).toBe(xml);
    expect(Boolean(textOf(reopened, 'Title').bold)).toBe(true);
  });

  test("clear 'all' drops the character style and keeps the paragraph style", async () => {
    const { xml, reopened } = await editSaveReopen(yrs, [
      format('00000004', { clear: 'all' }),
      format('00000003', { clear: 'all' }),
    ]);
    const strong = savedRunProperties(xml, 'Strong run') ?? '';
    expect(strong).not.toContain('w:rStyle');
    expect(strong).not.toContain('<w:b');
    expect(Boolean(textOf(reopened, 'Strong run').bold)).toBe(false);
    expect(Boolean(textOf(reopened, 'Title').bold)).toBe(true);
  });

  test('set and clear together save their marks', async () => {
    const { xml } = await editSaveReopen(yrs, [format('00000002', { clear: ['bold'], set: { italic: true, color: 'C00000' } })]);
    const properties = savedRunProperties(xml, 'Bold words');
    expect(holdsProperty(properties, 'w:i', {})).toBe(true);
    expect(holdsProperty(properties, 'w:color', { 'w:val': 'C00000' })).toBe(true);
    expect(properties ?? '').not.toContain('<w:b/>');
  });

  test('formatRange clear refuses a range across paragraphs and changes nothing', async () => {
    const session = await yrs.createYrsSession({ clientId: 73 });
    try {
      session.seedFromDocx(await textRunFixture());
      const before = session.encodeState();
      const across = { story: 'body', start: { paraId: '00000002', offset: 0 }, end: { paraId: '00000003', offset: 5 } };
      let code: unknown;
      try {
        session.formatRange(across, { clear: ['bold'] });
      } catch (error) {
        code = (error as { code?: unknown }).code;
      }
      expect(code).toBe('clear-across-paragraphs');
      expect(session.encodeState()).toEqual(before);
      const within = { story: 'body', start: { paraId: '00000003', offset: 0 }, end: { paraId: '00000003', offset: 5 } };
      session.formatRange(within, { clear: 'all', italic: true });
      const title = session.storySegments('body').find((entry) => entry.kind === 'text' && entry.text === 'Title');
      expect(title?.kind === 'text' && Boolean(title.attributes.bold) && Boolean(title.attributes.italic)).toBe(true);
    } finally {
      session.destroy();
    }
  });
});
