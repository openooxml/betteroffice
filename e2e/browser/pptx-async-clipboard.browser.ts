import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expectRepaint, inspectDeck, painted, slideCanvas, type SavedDeck } from './pptx-async-helpers';

const input = (page: Page) => page.getByRole('textbox', { name: 'Slide text', exact: true });
const slides = (page: Page) => page.getByRole('complementary', { name: 'Slides', exact: true }).getByRole('button');
const sourceText = 'Styled source';
const targetText = 'Target';

async function fixture() {
  const zip = new JSZip();
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  const shape = (id: number, text: string, y: number, rich = false) => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="914400" y="${y}"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"${rich ? ' b="1" i="1" u="sng"' : ''}>
    <a:solidFill><a:srgbClr val="${rich ? 'CC1122' : '172554'}"/></a:solidFill><a:latin typeface="Arial"/>
    </a:rPr><a:t>${text}</a:t></a:r><a:endParaRPr sz="2400"/></a:p></p:txBody></p:sp>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    ${[1, 2].map((index) => `<Override PartName="/ppt/slides/slide${index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
    <p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels([1, 2].map((index) => `<Relationship Id="rId${index}" Type="${r}/slide" Target="slides/slide${index}.xml"/>`).join('')));
  for (const index of [1, 2]) {
    zip.file(`ppt/slides/slide${index}.xml`, `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      ${index === 1 ? shape(2, sourceText, 914400, true) + shape(3, targetText, 2743200) : shape(2, targetText, 914400)}
      </p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, file?: string) {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles(file ?? {
    name: 'clipboard.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await fixture(),
  });
  await expect(slideCanvas(page)).toBeVisible();
  await expect(slides(page)).toHaveCount(2);
  await expect(slides(page).first()).toHaveAttribute('aria-current', 'page');
}

async function save(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${label}.pptx`);
  await (await downloading).saveAs(file);
  return { file, deck: await inspectDeck(page, await readFile(file)) };
}

async function selectText(page: Page, row = 0) {
  const canvas = await slideCanvas(page).boundingBox();
  if (!canvas) throw new Error('Slide canvas is missing');
  await page.mouse.click(canvas.x + canvas.width * 0.15, canvas.y + canvas.height * (row === 0 ? 0.22 : 0.58));
  await expect(input(page)).toBeFocused();
  await input(page).press('ControlOrMeta+Home');
  await input(page).press('ControlOrMeta+Shift+End');
  await expect(page.getByRole('status').filter({ hasText: 'Text selected' })).toBeVisible();
}

async function history(page: Page, info: TestInfo, before: SavedDeck, after: SavedDeck, originalPaint: string, row: number) {
  await page.getByTestId('pptx-undo').click();
  expect((await save(page, info, 'undo')).deck).toEqual(before);
  await selectText(page, row);
  await expect.poll(async () => (await painted(page)) === originalPaint).toBe(true);
  await page.getByTestId('pptx-redo').click();
  expect((await save(page, info, 'redo')).deck).toEqual(after);
  await expectRepaint(page, originalPaint);
}

async function reopen(page: Page, info: TestInfo, saved: Awaited<ReturnType<typeof save>>, slide = 0) {
  await open(page, saved.file);
  if (slide) await slides(page).nth(slide).click();
  await expect(slides(page).nth(slide)).toHaveAttribute('aria-current', 'page');
  expect((await save(page, info, 'reopened')).deck).toEqual(saved.deck);
}

test.beforeEach(async ({ context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
});

for (const multiline of [false, true]) {
  test(multiline ? 'external multiline paste creates separate paragraphs in one undo step' : 'external plain text paste replaces the selection in one undo step', async ({ page }, info) => {
    await open(page);
    const before = (await save(page, info, 'before')).deck;
    await selectText(page, 1);
    const originalPaint = await painted(page);
    const text = multiline ? 'First line\nSecond line\nThird line' : 'Outside text Ω';
    await page.evaluate((text) => navigator.clipboard.writeText(text), text);
    await input(page).press('ControlOrMeta+v');
    await expectRepaint(page, originalPaint);
    const saved = await save(page, info, 'pasted');
    expect(saved.deck[0]!.shapes[1]!.text).toBe(text);
    expect(saved.deck[0]!.shapes[1]!.paragraphs).toEqual(text.split('\n'));
    await history(page, info, before, saved.deck, originalPaint, 1);
    await reopen(page, info, saved);
    await selectText(page, 1);
    await expect(page.getByTestId('pptx-bold')).toHaveAttribute('aria-pressed', 'false');
  });
}

test('cut text copies the selection and undo restores it in one step', async ({ page }, info) => {
  await open(page);
  const before = (await save(page, info, 'before')).deck;
  await selectText(page);
  const originalPaint = await painted(page);
  await page.evaluate(() => navigator.clipboard.writeText('sentinel'));
  await input(page).press('ControlOrMeta+x');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(sourceText);
  await expectRepaint(page, originalPaint);
  const cut = await save(page, info, 'cut');
  expect(cut.deck[0]!.shapes[0]!.text).toBe('');
  await history(page, info, before, cut.deck, originalPaint, 0);
  await reopen(page, info, cut);
});
