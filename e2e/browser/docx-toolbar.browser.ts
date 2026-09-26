import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');

async function open(page: Page) {
  await page.goto('/?format=docx');
  await page
    .locator('input[type=file]')
    .last()
    .setInputFiles(resolve(root, 'apps/demo/public/betteroffice-demo.docx'));
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({
    timeout: 60_000,
  });
  await expect(page.getByTestId('formatting-bar')).toBeVisible();
}

async function savedText(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(label + '.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(
    (xml) =>
      Array.from(
        new DOMParser()
          .parseFromString(xml, 'application/xml')
          .getElementsByTagNameNS('*', 't')
      )
        .map((node) => node.textContent ?? '')
        .join(''),
    xml
  );
}

async function focusDocument(page: Page) {
  await page
    .locator('.canvas-page canvas')
    .first()
    .click({ position: { x: 150, y: 150 } });
  const input = page.getByTestId('yrs-input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  return input;
}

test('docx: a narrow toolbar keeps every control in a keyboard menu inside the viewport', async ({
  page,
}) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await open(page);
  const input = await focusDocument(page);
  await input.pressSequentially('Narrow');
  await input.press('Shift+Home');

  const more = page.getByTestId('formatting-bar').getByTestId('toolbar-more');
  await expect(more).toBeVisible();
  const scrollWidth = () =>
    page.evaluate(() => document.documentElement.scrollWidth);
  const pageWidth = await scrollWidth();
  await more.focus();
  await page.keyboard.press('ArrowDown');
  const menu = page.getByRole('menu', { name: 'More actions' });
  await expect(menu).toBeVisible();
  const box = (await menu.boundingBox())!;
  expect(box.x).toBeGreaterThanOrEqual(0);
  expect(box.x + box.width).toBeLessThanOrEqual(360);
  expect(await scrollWidth()).toBe(pageWidth);

  await page.keyboard.press('End');
  await expect(page.locator(':focus')).toHaveAttribute(
    'data-label',
    'Editing mode'
  );
  await page.keyboard.press('Home');
  await page.keyboard.type('b');
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Bold');
  await page.keyboard.press('Enter');
  await expect(menu).toBeHidden();
  await expect(more).toBeFocused();

  await page.keyboard.press('ArrowDown');
  const bold = menu.locator('[data-label="Bold"]');
  await expect(bold).toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(more).toBeFocused();
});

const LONGER = ' (übersetzt, mit deutlich längerer Beschriftung)';

type Strings = { [key: string]: string | Strings };

async function englishStrings(): Promise<Strings> {
  return JSON.parse(
    await readFile(resolve(root, 'packages/docx-i18n/en.json'), 'utf8')
  ) as Strings;
}

function lookup(strings: Strings, key: string): string {
  const value = key
    .split('.')
    .reduce<string | Strings>((node, part) => (node as Strings)[part], strings);
  if (typeof value !== 'string') throw new Error(`No string at ${key}`);
  return value;
}

async function lengthenStrings(page: Page, strings: Strings) {
  const pairs: [string, string][] = [];
  const walk = (node: Strings) => {
    for (const [key, value] of Object.entries(node)) {
      if (typeof value === 'object') walk(value);
      else if (!key.startsWith('_')) pairs.push([key, value]);
    }
  };
  walk(strings);
  const marker = /highlightColor:[`"']Text Highlight Color[`"']/;
  await page.route('**/assets/*.js', async (route) => {
    const response = await route.fetch();
    let body = await response.text();
    if (marker.test(body)) {
      for (const [key, value] of pairs) {
        for (const quote of ['`', '"', "'"]) {
          body = body
            .split(`${key}:${quote}${value}${quote}`)
            .join(`${key}:${quote}${value}${LONGER}${quote}`);
        }
      }
    }
    await route.fulfill({ response, body });
  });
}

const DEFAULT_TOOLBAR: [group: string | null, controls: string[]][] = [
  [
    'formattingBar.groups.history',
    ['formattingBar.undo', 'formattingBar.redo'],
  ],
  ['formattingBar.groups.zoom', ['zoom.zoomLevel']],
  ['formattingBar.groups.styles', ['commands.paragraphStyle']],
  ['formattingBar.groups.font', ['commands.fontFamily', 'fontSize.label']],
  [
    'formattingBar.groups.textFormatting',
    [
      'formattingBar.bold',
      'formattingBar.italic',
      'formattingBar.underline',
      'formattingBar.strikethrough',
      'formattingBar.fontColor',
      'formattingBar.highlightColor',
      'formattingBar.insertLink',
    ],
  ],
  [
    'formattingBar.groups.script',
    ['formattingBar.superscript', 'formattingBar.subscript'],
  ],
  ['formattingBar.groups.alignment', ['formattingBar.groups.alignment']],
  [
    'formattingBar.groups.listFormatting',
    [
      'lists.bulletList',
      'lists.numberedList',
      'lists.decreaseIndent',
      'lists.increaseIndent',
      'lineSpacing.label',
    ],
  ],
  [null, ['formattingBar.clearFormatting']],
  [null, ['editor.toggleCommentsSidebar']],
  [null, ['commands.editingMode']],
];

test('docx: long translations keep every toolbar action reachable through More', async ({
  page,
}) => {
  const strings = await englishStrings();
  const long = (key: string) => lookup(strings, key) + LONGER;
  await lengthenStrings(page, strings);
  await page.setViewportSize({ width: 360, height: 800 });
  await open(page);
  await focusDocument(page);

  const rail = page.getByTestId('formatting-bar');
  const more = rail.getByTestId('toolbar-more');
  await expect(more).toHaveAttribute(
    'aria-label',
    long('commands.moreActions')
  );
  await more.focus();
  await page.keyboard.press('ArrowDown');
  const menu = page.getByRole('menu', { name: long('commands.moreActions') });
  await expect(menu).toBeVisible();

  const row = rail.locator('[data-toolbar-items]');
  let overflowed = 0;
  for (const [group, controls] of DEFAULT_TOOLBAR) {
    if (
      group &&
      (await row.getByRole('group', { name: long(group), exact: true }).count())
    ) {
      continue;
    }
    for (const control of controls) {
      await expect(menu.locator(`[data-label="${long(control)}"]`)).toHaveCount(
        1
      );
      overflowed += 1;
    }
  }
  expect(overflowed).toBeGreaterThan(20);

  const fits = (locator: typeof menu) =>
    locator.evaluate((element) => {
      const box = element.getBoundingClientRect();
      const items = Array.from(
        element.querySelectorAll<HTMLElement>('[role^="menuitem"]')
      );
      return (
        box.left >= 0 &&
        box.right <= window.innerWidth &&
        items.every((item) => item.scrollWidth <= item.clientWidth + 1)
      );
    });
  expect(await fits(menu)).toBe(true);

  const color = menu.locator(
    `[data-label="${long('formattingBar.fontColor')}"]`
  );
  await color.focus();
  await page.keyboard.press('ArrowRight');
  const colors = page.getByRole('menu', {
    name: long('formattingBar.fontColor'),
  });
  await expect(colors).toBeVisible();
  await expect(
    colors.locator(`[data-label="${long('colorPicker.automatic')}"]`)
  ).toHaveCount(1);
  expect(await fits(colors)).toBe(true);
});

test('docx: a toolbar command issued during IME composition applies after the committed text', async ({
  page,
}, info) => {
  await open(page);
  const input = await focusDocument(page);
  await input.pressSequentially('ab');
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', {
    text: 'かな',
    selectionStart: 2,
    selectionEnd: 2,
  });
  const rail = page.getByTestId('formatting-bar');
  const undo = rail.getByRole('button', { name: 'Undo' });
  await expect(undo).not.toHaveAttribute('aria-disabled', 'true');
  await undo.click();
  await cdp.send('Input.insertText', { text: '漢字' });

  const redo = rail.getByRole('button', { name: 'Redo' });
  await expect(redo).not.toHaveAttribute('aria-disabled', 'true');
  expect(await savedText(page, info, 'undone')).not.toContain('漢字');
  await redo.click();
  expect(await savedText(page, info, 'redone')).toContain('漢字');
});
