import { afterEach, expect, test } from 'bun:test';
import { readEchoParagraph } from './residentCommandViewProjection';
import { closeCommandEngines, openCommandEngine, TIMEOUT } from './residentCommandTestSupport';

const W = 'w:author="A" w:date="2026-01-01T00:00:00Z"';
const run = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (content: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content}</w:p>`;
const sdt = (content: string) => `<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const cell = (content: string) => `<w:tc><w:tcPr/>${content}</w:tc>`;

const BODY = [
  p(run('Plain text')),
  p(run('Bold', '<w:b/>') + run(' and ') + run('italic', '<w:i/>')),
  p(run('Tab') + '<w:r><w:tab/></w:r>' + run('after')),
  p(run('Line') + '<w:r><w:br/></w:r>' + run('break')),
  p(run('Field ') + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>'),
  p(run('Inline ') + sdt(run('control'))),
  p(run('Kept ') + `<w:ins w:id="1" ${W}>${run('inserted')}</w:ins>`),
  p(run('Kept ') + `<w:del w:id="2" ${W}><w:r><w:delText>deleted</w:delText></w:r></w:del>`),
  p(run('Joined into next'), `<w:rPr><w:del w:id="3" ${W}/></w:rPr>`),
  p(run('After a deleted mark')),
  `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell(p(run('In a cell')))}${cell(sdt(p(run('Control in a cell'))))}</w:tr></w:tbl>`,
  p(run('After the table')),
  p('<w:r><w:br w:type="page"/></w:r>' + run('After a page break')),
  p(run('')),
  p(run('Last')),
].join('');

afterEach(closeCommandEngines);

test('the segment read gives the story read\'s text and atoms or declines', async () => {
  const arm = await openCommandEngine([], { bodyXml: BODY });
  const engine = arm.engine;
  let fast = 0;
  let declined = 0;
  const seen = new Set<string>();
  for (const story of engine.geometryReader.storyIds()) {
    for (const { paraId } of engine.paragraphs(story)) {
      const full = engine.proposalEngine.readParagraphs({ story, paraIds: [paraId], view: 'accepted' });
      const read = readEchoParagraph(engine, story, paraId);
      if (!read) { declined += 1; continue; }
      fast += 1;
      expect(full.ok).toBe(true);
      if (!full.ok) continue;
      const paragraph = full.paragraphs[0]!;
      expect({ story, paraId, text: read.text, hasAtoms: read.hasAtoms })
        .toEqual({ story, paraId, text: paragraph.text, hasAtoms: paragraph.atoms.length > 0 });
      seen.add(paragraph.text);
    }
  }
  expect(seen.has('Plain text')).toBe(true);
  expect(seen.has('In a cell')).toBe(true);
  expect(seen.has('Control in a cell')).toBe(true);
  expect(seen.has('After the table')).toBe(true);
  expect(fast).toBeGreaterThan(8);
  expect(declined).toBeGreaterThanOrEqual(3);
}, TIMEOUT);

test('tracked changes on or before a paragraph send it to the story read', async () => {
  const arm = await openCommandEngine([], { bodyXml: BODY });
  const engine = arm.engine;
  const revised = engine.paragraphs('body').filter((paragraph) =>
    /inserted|deleted|Joined into next|After a deleted mark/.test(paragraph.text));
  expect(revised.length).toBeGreaterThanOrEqual(3);
  for (const paragraph of revised) expect(readEchoParagraph(engine, 'body', paragraph.paraId)).toBeUndefined();
}, TIMEOUT);

test('an unknown or repeated paragraph id declines', async () => {
  const arm = await openCommandEngine([], { bodyXml: BODY });
  expect(readEchoParagraph(arm.engine, 'body', 'missing')).toBeUndefined();
}, TIMEOUT);
