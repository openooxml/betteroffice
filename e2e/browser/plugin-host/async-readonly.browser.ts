import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { pinnedEvaluate } from '../pinned-evaluate';
import type {} from './xlsx-worker-editor-probe';
import type {} from './pptx-peerless-editor-probe';
import { renderedBox } from '../rendered-box';

type Format = 'xlsx' | 'pptx';
const errors = new WeakMap<Page, string[]>();
const slideInput = (page: Page) => page.getByRole('textbox', { name: 'Slide text', exact: true });

async function open(page: Page, format: 'xlsx' | 'pptx') {
  await page.goto(`/${format}-worker-editor.html?asyncApi=1`);
  if (format === 'xlsx') {
    await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
    await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.ready);
    await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
  } else {
    await page.waitForFunction(() => '__pptxPeerlessEditor' in window);
    await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.ready);
    await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('initial');
  }
}

async function readOnly(page: Page, format: 'xlsx' | 'pptx', value = true, toolbar = true) {
  await page.evaluate(({ format, value, toolbar }) => {
    if (format === 'xlsx') window.__xlsxWorkerEditor.setReadOnly(value, toolbar);
    else window.__pptxPeerlessEditor.setReadOnly(value, toolbar);
  }, { format, value, toolbar });
  if (format === 'pptx') {
    await expect(page.getByRole('application')).toHaveAttribute('aria-readonly', String(value));
    if (value) await expect(slideInput(page)).not.toBeEditable();
    else await expect(slideInput(page)).toBeEditable();
  } else if (toolbar) {
    if (value) await expect(page.getByTestId('xlsx-formula-input')).not.toBeEditable();
    else await expect(page.getByTestId('xlsx-formula-input')).toBeEditable();
  } else {
    await expect(page.getByTestId('xlsx-toolbar')).toHaveCount(value ? 0 : 1);
  }
}

async function save(page: Page, format: Format): Promise<Buffer> {
  const count = await page.evaluate((format) => format === 'xlsx'
    ? window.__xlsxWorkerEditor.saves.length : window.__pptxPeerlessEditor.saves.length, format);
  if (format === 'xlsx') await page.getByTestId('xlsx-scroll').press('ControlOrMeta+s');
  else await slideInput(page).press('ControlOrMeta+s');
  await expect.poll(() => page.evaluate((format) => format === 'xlsx'
    ? window.__xlsxWorkerEditor.saves.length : window.__pptxPeerlessEditor.saves.length, format)).toBe(count + 1);
  return Buffer.from(await page.evaluate((format) => (format === 'xlsx'
    ? window.__xlsxWorkerEditor.saves : window.__pptxPeerlessEditor.saves).at(-1)!, format));
}

async function xml(bytes: Buffer, path: string) {
  const part = (await JSZip.loadAsync(bytes)).file(path);
  if (!part) throw new Error(`Saved document is missing ${path}`);
  return part.async('string');
}

async function savedText(page: Page, bytes: Buffer, format: Format) {
  const zip = await JSZip.loadAsync(bytes);
  const paths = format === 'xlsx'
    ? Object.keys(zip.files).filter((path) => /^xl\/(worksheets\/sheet\d+|sharedStrings)\.xml$/.test(path)).sort()
    : ['ppt/slides/slide1.xml'];
  const parts = await Promise.all(paths.map((path) => zip.file(path)!.async('string')));
  return page.evaluate((parts) => parts.map((source) => {
    const document = new DOMParser().parseFromString(source, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Saved XML is invalid');
    return Array.from(document.getElementsByTagNameNS('*', 't')).map((node) => node.textContent ?? '').join('');
  }).join(''), parts);
}

async function unchanged(page: Page, format: Format, baseline: Buffer) {
  const bytes = await save(page, format);
  expect(bytes.equals(baseline), 'Read-only save must preserve every byte').toBe(true);
  expect(await savedText(page, bytes, format)).toBe(await savedText(page, baseline, format));
}

async function clipboard(page: Page, text: string) {
  await pinnedEvaluate(page, (text) => navigator.clipboard.writeText(text), text);
}

async function copied(page: Page, text: string) {
  await expect.poll(() => pinnedEvaluate(page, () => navigator.clipboard.readText())).toBe(text);
}

async function xlsxInitial(page: Page, text = 'initial') {
  await expect.poll(() => pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1')))
    .toEqual([[text, '1']]);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain(text);
}

async function editCell(page: Page, value: string) {
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.press('ControlOrMeta+Home');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await scroll.press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await expect(input).toHaveValue(/.+/);
  await input.press('ControlOrMeta+A');
  await input.pressSequentially(value);
  await input.press('Enter');
  await expect(input).toHaveCount(0);
  await xlsxInitial(page, value);
}

async function selectSlideText(page: Page) {
  const point = await page.evaluate(() => window.__pptxPeerlessEditor.point(0));
  await page.mouse.click(point.x, point.y);
  await expect(slideInput(page)).toBeFocused();
  await slideInput(page).press('Home');
  await slideInput(page).press('Shift+End');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.selection()))
    .toMatchObject({ anchor: 0, focus: 7 });
}

async function pptxInitial(page: Page, text = 'initial') {
  await expect.poll(async () => (await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.content())).stories[0]!.text).toBe(text);
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain(text);
}

test.beforeEach(async ({ page, context }) => {
  const messages: string[] = [];
  errors.set(page, messages);
  page.on('pageerror', (error) => messages.push(error.message));
  page.on('console', (message) => { if (message.type() === 'error') messages.push(message.text()); });
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
});

test.afterEach(async ({ page }) => {
  expect.soft(errors.get(page), 'Console and page errors').toEqual([]);
  const probes = await page.evaluate(() => ({
    xlsx: window.__xlsxWorkerEditor ? { engines: window.__xlsxWorkerEditor.mainThreadEngines(), errors: window.__xlsxWorkerEditor.errors } : null,
    pptx: window.__pptxPeerlessEditor ? { engines: window.__pptxPeerlessEditor.mainThreadEngines(), errors: window.__pptxPeerlessEditor.errors } : null,
  }));
  for (const probe of [probes.xlsx, probes.pptx]) {
    if (!probe) continue;
    expect.soft(probe.engines, 'Main-thread document engines').toBe(0);
    expect.soft(probe.errors, 'Editor errors').toEqual([]);
  }
});

test('xlsx read-only refuses typing, paste, cut and deletion', async ({ page }) => {
  await open(page, 'xlsx');
  const baseline = await save(page, 'xlsx');
  await readOnly(page, 'xlsx', true, false);
  const scroll = page.getByTestId('xlsx-scroll');
  await clipboard(page, 'blocked paste');
  for (const key of ['x', 'ControlOrMeta+v', 'ControlOrMeta+x', 'Delete', 'Backspace']) {
    await scroll.press(key);
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
    await xlsxInitial(page);
    await unchanged(page, 'xlsx', baseline);
  }
});

test('xlsx read-only refuses F2, double-click and formula bar edits', async ({ page }) => {
  await open(page, 'xlsx');
  const baseline = await save(page, 'xlsx');
  await readOnly(page, 'xlsx');
  await page.getByTestId('xlsx-scroll').press('F2');
  await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
  await unchanged(page, 'xlsx', baseline);
  const rect = await page.evaluate(() => window.__xlsxWorkerEditor.cellClip());
  await page.mouse.dblclick(rect.x + rect.width / 2, rect.y + rect.height / 2);
  await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
  await unchanged(page, 'xlsx', baseline);
  const formula = page.getByTestId('xlsx-formula-input');
  await expect(formula).toHaveValue('initial');
  await formula.press('ControlOrMeta+A');
  await formula.pressSequentially('blocked formula');
  await formula.press('Enter');
  await expect(formula).toHaveValue('initial');
  await xlsxInitial(page);
  await unchanged(page, 'xlsx', baseline);
});

test('xlsx read-only refuses formatting shortcuts and toolbar mutations', async ({ page }) => {
  await open(page, 'xlsx');
  const baseline = await save(page, 'xlsx');
  await readOnly(page, 'xlsx');
  for (const key of ['ControlOrMeta+b', 'ControlOrMeta+i']) {
    await page.getByTestId('xlsx-scroll').press(key);
    await unchanged(page, 'xlsx', baseline);
  }
  for (const name of ['Bold', 'Italic']) {
    const button = page.getByRole('button', { name, exact: true });
    await expect(button).toBeDisabled();
    await expect(button).toHaveAttribute('aria-pressed', 'false');
    await button.click({ force: true });
    await unchanged(page, 'xlsx', baseline);
  }
  expect(await xml(await save(page, 'xlsx'), 'xl/styles.xml')).toBe(await xml(baseline, 'xl/styles.xml'));
  await xlsxInitial(page);
});

test('xlsx read-only preserves selection, copy, scrolling, zoom and sheet navigation', async ({ page }) => {
  await open(page, 'xlsx');
  await clipboard(page, Array.from({ length: 80 }, (_, index) => index === 0 ? 'initial\t1' : index === 79 ? 'scroll end' : '').join('\n'));
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+v');
  await expect.poll(() => pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.cells(0, 'A80'))).toEqual([['scroll end']]);
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+Home');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  const baseline = await save(page, 'xlsx');
  await readOnly(page, 'xlsx');
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.press('ArrowRight');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B1');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('1');
  await scroll.press('Shift+ArrowLeft');
  await expect.poll(() => pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.selection()))
    .toMatchObject({ sheetIndex: 0, cells: { anchor: { row: 0, col: 1 }, focus: { row: 0, col: 0 } } });
  await clipboard(page, 'copy sentinel');
  await scroll.press('ControlOrMeta+c');
  await copied(page, 'initial\t1');
  await scroll.hover({ position: { x: 300, y: 300 } });
  await page.mouse.wheel(0, 500);
  await expect.poll(() => scroll.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  await scroll.press('ControlOrMeta+Home');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect.poll(() => scroll.evaluate((element) => element.scrollTop)).toBe(0);
  const zoom = page.getByRole('combobox', { name: /^Zoom:/ });
  await zoom.fill('125%');
  await zoom.press('Enter');
  await expect(page.locator('[data-paint-source="worker"]')).toHaveAttribute('data-worker-zoom', '1.25');
  await page.getByRole('tab', { name: 'Other', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'Other', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('other');
  await page.getByRole('tab', { name: 'Sheet', exact: true }).click();
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('initial');
  await unchanged(page, 'xlsx', baseline);
});

test('xlsx edits survive read-only toggles, history, save and reopen', async ({ page }) => {
  await open(page, 'xlsx');
  await editCell(page, 'retained edit');
  const baseline = await save(page, 'xlsx');
  expect(await savedText(page, baseline, 'xlsx')).toContain('retained edit');
  const generation = await page.evaluate(() => window.__xlsxWorkerEditor.generation());
  await readOnly(page, 'xlsx');
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+z');
  await unchanged(page, 'xlsx', baseline);
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+Shift+z');
  await unchanged(page, 'xlsx', baseline);
  await xlsxInitial(page, 'retained edit');
  await readOnly(page, 'xlsx', false);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.generation())).toBe(generation);
  await editCell(page, 'editing again');
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+z');
  await xlsxInitial(page, 'retained edit');
  expect(await savedText(page, await save(page, 'xlsx'), 'xlsx')).toContain('retained edit');
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+Shift+z');
  await xlsxInitial(page, 'editing again');
  expect(await savedText(page, await save(page, 'xlsx'), 'xlsx')).toContain('editing again');
  await readOnly(page, 'xlsx');
  await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.saveAndReopen());
  await xlsxInitial(page, 'editing again');
  expect(await savedText(page, await save(page, 'xlsx'), 'xlsx')).toContain('editing again');
});

test('pptx read-only refuses typing, paste, cut and deletion', async ({ page }) => {
  await open(page, 'pptx');
  const baseline = await save(page, 'pptx');
  await readOnly(page, 'pptx', true, false);
  await selectSlideText(page);
  await clipboard(page, 'blocked paste');
  for (const key of ['x', 'ControlOrMeta+v', 'ControlOrMeta+x', 'Delete', 'Backspace']) {
    await slideInput(page).press(key);
    await pptxInitial(page);
    await unchanged(page, 'pptx', baseline);
  }
});

test('pptx read-only refuses formatting shortcuts and toolbar mutations', async ({ page }) => {
  await open(page, 'pptx');
  const baseline = await save(page, 'pptx');
  await readOnly(page, 'pptx');
  await selectSlideText(page);
  for (const key of ['ControlOrMeta+b', 'ControlOrMeta+i', 'ControlOrMeta+u']) {
    await slideInput(page).press(key);
    await unchanged(page, 'pptx', baseline);
  }
  for (const name of ['Bold', 'Italic', 'Underline']) {
    const button = page.getByRole('button', { name, exact: true });
    await expect(button).toBeDisabled();
    await button.click({ force: true });
    await unchanged(page, 'pptx', baseline);
  }
  for (const name of ['Text box', 'Insert image']) {
    const button = page.getByRole('button', { name, exact: true });
    await expect(button).toBeDisabled();
    await button.click({ force: true });
    await unchanged(page, 'pptx', baseline);
  }
  await pptxInitial(page);
});

test('pptx read-only refuses shape drag, resize, nudge and removal', async ({ page }) => {
  await open(page, 'pptx');
  const point = await page.evaluate(() => window.__pptxPeerlessEditor.shapePoint());
  await page.mouse.click(point.x, point.y);
  await expect(page.getByTestId('pptx-shape-selection')).toBeVisible();
  await expect(page.getByTestId('pptx-resize-se')).toBeVisible();
  const baseline = await save(page, 'pptx');
  const original = (await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.content())).slides[0]!.shapes;
  await readOnly(page, 'pptx');
  await expect(page.getByTestId('pptx-resize-se')).toHaveCount(0);
  const shape = await page.evaluate(() => window.__pptxPeerlessEditor.shapePoint());
  await page.mouse.move(shape.x, shape.y);
  await page.mouse.down();
  await page.mouse.move(shape.x + 40, shape.y + 30, { steps: 4 });
  await page.mouse.up();
  await unchanged(page, 'pptx', baseline);
  const selection = await renderedBox(page.getByTestId('pptx-shape-selection'));
  await page.mouse.move(selection.x + selection.width, selection.y + selection.height);
  await page.mouse.down();
  await page.mouse.move(selection.x + selection.width + 40, selection.y + selection.height + 30, { steps: 4 });
  await page.mouse.up();
  await unchanged(page, 'pptx', baseline);
  await page.mouse.click(shape.x, shape.y);
  const stage = page.getByRole('application');
  for (const key of ['ArrowRight', 'Shift+ArrowDown', 'Delete']) {
    await stage.press(key);
    await unchanged(page, 'pptx', baseline);
  }
  expect((await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.content())).slides[0]!.shapes).toEqual(original);
  await pptxInitial(page);
});

test('pptx read-only refuses slide insertion, deletion and reordering', async ({ page }) => {
  await open(page, 'pptx');
  const baseline = await save(page, 'pptx');
  await readOnly(page, 'pptx');
  const insert = page.getByRole('button', { name: 'New slide', exact: true });
  await expect(insert).toBeDisabled();
  await insert.click({ force: true });
  await expect(page.getByRole('button', { name: 'Add slide', exact: true })).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Delete slide', exact: true })).toHaveCount(0);
  await unchanged(page, 'pptx', baseline);
  const slides = page.getByRole('complementary').getByRole('button');
  await expect(slides).toHaveCount(2);
  await expect(slides.first()).toHaveAttribute('draggable', 'false');
  await slides.first().press('Delete');
  await unchanged(page, 'pptx', baseline);
  await slides.first().press('Alt+ArrowDown');
  await unchanged(page, 'pptx', baseline);
  await expect(slides).toHaveCount(2);
  expect((await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.content())).stories.map((story) => story.text)).toEqual(['initial', 'other']);
});

test('pptx read-only preserves copy, slide navigation, zoom and scrolling', async ({ page }) => {
  await open(page, 'pptx');
  const baseline = await save(page, 'pptx');
  await readOnly(page, 'pptx');
  await selectSlideText(page);
  await clipboard(page, 'copy sentinel');
  await slideInput(page).press('ControlOrMeta+c');
  await copied(page, 'initial');
  const slides = page.getByRole('complementary').getByRole('button');
  await slides.nth(1).click();
  await expect(slides.nth(1)).toHaveAttribute('aria-current', 'page');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('other');
  const stage = page.getByRole('application');
  await stage.press('Home');
  await expect(slides.first()).toHaveAttribute('aria-current', 'page');
  await stage.press('PageDown');
  await expect(slides.nth(1)).toHaveAttribute('aria-current', 'page');
  await stage.press('PageUp');
  await expect(slides.first()).toHaveAttribute('aria-current', 'page');
  const zoom = page.getByRole('combobox', { name: /^Zoom:/ });
  await zoom.fill('200%');
  await zoom.press('Enter');
  await expect(zoom).toHaveValue('200%');
  await stage.hover({ position: { x: 300, y: 300 } });
  await page.mouse.wheel(0, 400);
  await expect.poll(() => stage.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  await unchanged(page, 'pptx', baseline);
});

test('pptx edits survive read-only toggles, history, save and reopen', async ({ page }) => {
  await open(page, 'pptx');
  await selectSlideText(page);
  await slideInput(page).pressSequentially('retained edit');
  await pptxInitial(page, 'retained edit');
  const baseline = await save(page, 'pptx');
  expect(await savedText(page, baseline, 'pptx')).toBe('retained edit');
  const generation = await page.evaluate(() => window.__pptxPeerlessEditor.generation());
  await readOnly(page, 'pptx');
  await slideInput(page).press('ControlOrMeta+z');
  await unchanged(page, 'pptx', baseline);
  await slideInput(page).press('ControlOrMeta+Shift+z');
  await unchanged(page, 'pptx', baseline);
  await readOnly(page, 'pptx', false);
  expect(await page.evaluate(() => window.__pptxPeerlessEditor.generation())).toBe(generation);
  const point = await page.evaluate(() => window.__pptxPeerlessEditor.point(0));
  await page.mouse.click(point.x, point.y);
  await slideInput(page).press('End');
  await clipboard(page, ' again');
  await slideInput(page).press('ControlOrMeta+v');
  await pptxInitial(page, 'retained edit again');
  await slideInput(page).press('ControlOrMeta+z');
  await pptxInitial(page, 'retained edit');
  expect(await savedText(page, await save(page, 'pptx'), 'pptx')).toBe('retained edit');
  await slideInput(page).press('ControlOrMeta+Shift+z');
  await pptxInitial(page, 'retained edit again');
  expect(await savedText(page, await save(page, 'pptx'), 'pptx')).toBe('retained edit again');
  await readOnly(page, 'pptx');
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.saveAndReopen());
  await pptxInitial(page, 'retained edit again');
  expect(await savedText(page, await save(page, 'pptx'), 'pptx')).toBe('retained edit again');
});
