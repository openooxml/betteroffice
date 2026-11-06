import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

type Probe = {
  issues: { code: string; id: string }[];
  hostKeys: { key: string; prevented: boolean }[];
  saves: number;
  pluginRuns: number;
  load(bytes: number[], shortcuts: unknown): Promise<void>;
  state(id: 'bold' | 'italic'): boolean | 'mixed' | null | undefined;
  texts(): Promise<string[]>;
  selectTarget(): Promise<boolean>;
};

const probe = <T>(page: Page, pick: string) =>
  page.evaluate((body) => new Function('probe', `return ${body}`)(
    (window as unknown as { __shortcuts: Probe }).__shortcuts), pick) as Promise<T>;
const lastKey = (page: Page) => probe<{ key: string; prevented: boolean } | undefined>(page, 'probe.hostKeys.at(-1)');
const canvas = (page: Page) => page.getByTestId('pptx-slide-canvas');
const input = (page: Page) => page.getByRole('textbox', { name: 'Slide text', exact: true });

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
      ${index === 1 ? shape(2, 'First box', 914400, true) + shape(3, 'Target', 2743200) : shape(2, 'Target', 914400)}
      </p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, mode: string, shortcuts: unknown) {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto(`/pptx-shortcuts.html?mode=${mode}`);
  const bytes = [...await fixture()];
  await page.evaluate(([data, map]) => (window as unknown as { __shortcuts: Probe }).__shortcuts.load(data as number[], map),
    [bytes, shortcuts] as const);
  await expect(canvas(page)).toBeVisible({ timeout: 120_000 });
  await expect.poll(() => probe(page, 'probe.editor !== null'), { timeout: 60_000 }).toBe(true);
}

async function selectTarget(page: Page) {
  await expect(canvas(page)).toHaveAttribute('data-pptx-input-ready', 'true');
  expect(await probe<boolean>(page, 'probe.selectTarget()')).toBe(true);
  await input(page).focus();
  await expect(input(page)).toBeFocused();
}

test.describe('pptx host shortcuts, worker', () => {
  test('a remapped paste leaves the old chord to the host', async ({ page }) => {
    await open(page, 'worker', { paste: 'Mod+Alt+V' });
    await selectTarget(page);
    await page.evaluate(() => navigator.clipboard.writeText('Outside'));
    await input(page).press('Control+V');
    await page.waitForTimeout(400);
    expect(await probe<string[]>(page, 'probe.texts()')).not.toContain('Outside');
    expect(await lastKey(page)).toEqual({ key: 'v', prevented: false });
    await input(page).press('Control+Alt+V');
    await expect.poll(() => probe<string[]>(page, 'probe.texts()')).toContain('Outside');
  });

  test('a released plain paste chord pastes nothing; the default one pastes', async ({ page }) => {
    await open(page, 'worker', { pasteAsPlainText: null });
    await selectTarget(page);
    await page.evaluate(() => navigator.clipboard.writeText('Plain'));
    await input(page).press('Control+Shift+V');
    await page.waitForTimeout(400);
    expect(await probe<string[]>(page, 'probe.texts()')).not.toContain('Plain');
    expect(await lastKey(page)).toEqual({ key: 'v', prevented: false });
    await input(page).press('Control+V');
    await expect.poll(() => probe<string[]>(page, 'probe.texts()')).toContain('Plain');
  });

  test('moving bold frees its chord for another command', async ({ page }) => {
    await open(page, 'worker', { bold: 'Mod+Alt+B', italic: 'Mod+B' });
    expect(await probe(page, 'probe.issues')).toEqual([]);
    await selectTarget(page);
    await input(page).press('Control+B');
    await expect.poll(() => probe(page, 'probe.state("italic")')).toBe(true);
    expect(await probe(page, 'probe.state("bold")')).toBe(false);
    await input(page).press('Control+Alt+B');
    await expect.poll(() => probe(page, 'probe.state("bold")')).toBe(true);
  });

  test('a released copy chord is left to the host', async ({ page }) => {
    await open(page, 'worker', { copy: null, cut: 'Mod+Alt+X' });
    await selectTarget(page);
    await page.evaluate(() => navigator.clipboard.writeText('sentinel'));
    await input(page).press('Control+C');
    await page.waitForTimeout(400);
    expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('sentinel');
    await input(page).press('Control+Shift+X');
    await page.waitForTimeout(400);
    expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('sentinel');
    expect((await probe<string[]>(page, 'probe.texts()')).filter((text) => text === 'Target')).toHaveLength(2);
    await input(page).press('Control+Alt+X');
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Target');
  });

  test('the context menu shows the remapped chord', async ({ page }) => {
    await open(page, 'worker', { paste: 'Mod+Alt+V', copy: null });
    await selectTarget(page);
    await input(page).press('Shift+F10');
    const paste = page.getByRole('menuitem', { name: 'Paste', exact: true });
    await expect(paste).toContainText('Ctrl+Alt+V');
    await expect(paste).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+V');
    await expect(page.getByRole('menuitem', { name: 'Copy', exact: true })).not.toContainText('Ctrl');
  });
});

for (const mode of ['worker', 'legacy']) {
  test.describe(`pptx host shortcuts, ${mode}`, () => {
    test('a moved save chord runs; the old one reaches the host', async ({ page }) => {
      await open(page, mode, { save: 'Mod+Alt+S', bold: 'Mod+Alt+B' });
      await canvas(page).click();
      await page.keyboard.press('Control+S');
      await page.waitForTimeout(300);
      expect(await probe(page, 'probe.saves')).toBe(0);
      expect(await lastKey(page)).toEqual({ key: 's', prevented: false });
      await page.keyboard.press('Control+Alt+S');
      await expect.poll(() => probe(page, 'probe.saves')).toBe(1);
      const bold = page.getByTestId('pptx-bold');
      await expect(bold).toHaveAttribute('title', /Ctrl\+Alt\+B/);
      await expect(bold).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+B');
    });

    test('clipboard ids are refused only where there is no clipboard', async ({ page }) => {
      await open(page, mode, { paste: 'Mod+Alt+V' });
      const issues = await probe<{ code: string; id: string }[]>(page, 'probe.issues');
      expect(issues.map(({ code, id }) => [code, id])).toEqual(mode === 'legacy' ? [['unknown-id', 'paste']] : []);
    });

    test('refused entries are reported and plugin chords follow the map', async ({ page }) => {
      await open(page, mode, { bold: 'Mod+ArrowLeft', nope: 'Mod+J', 'plugin:acme.keys/run': 'Mod+Alt+K' });
      expect((await probe<{ code: string; id: string }[]>(page, 'probe.issues')).map(({ code, id }) => [code, id])).toEqual([
        ['reserved-chord', 'bold'],
        ['unknown-id', 'nope'],
      ]);
      await canvas(page).click();
      await page.keyboard.press('Control+Shift+K');
      await page.waitForTimeout(300);
      expect(await probe(page, 'probe.pluginRuns')).toBe(0);
      await page.keyboard.press('Control+Alt+K');
      await expect.poll(() => probe(page, 'probe.pluginRuns')).toBe(1);
    });
  });
}
