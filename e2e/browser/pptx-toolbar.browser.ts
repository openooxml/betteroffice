import { test, expect, type Locator, type Page } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');

async function open(page: Page) {
  await page.goto('/?format=pptx');
  await page
    .locator('input[type=file]')
    .last()
    .setInputFiles(resolve(root, 'apps/demo/public/betteroffice-demo.pptx'));
  await expect(page.getByTestId('pptx-slide-canvas')).toBeVisible({
    timeout: 60_000,
  });
  await expect(page.getByTestId('pptx-formatting-toolbar')).toBeVisible();
}

const fits = (locator: Locator) =>
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

test('pptx: a narrow toolbar keeps every control in a keyboard menu inside the viewport', async ({
  page,
}) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await open(page);
  const scrollWidth = () =>
    page.evaluate(() => document.documentElement.scrollWidth);
  const pageWidth = await scrollWidth();
  const rail = page.getByTestId('pptx-formatting-toolbar');
  const more = rail.getByTestId('pptx-toolbar-more');
  await expect(more).toBeVisible();

  await more.focus();
  await page.keyboard.press('ArrowDown');
  const menu = page.getByRole('menu', { name: 'More' });
  await expect(menu).toBeVisible();
  expect(await fits(menu)).toBe(true);
  expect(await scrollWidth()).toBe(pageWidth);

  await page.keyboard.press('End');
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Arrange');
  await page.keyboard.press('Home');
  await page.keyboard.type('z');
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Zoom');
  await page.keyboard.press('ArrowRight');
  const zoom = page.getByRole('menu', { name: 'Zoom' });
  await expect(zoom).toBeVisible();
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Fit');
  await expect(page.locator(':focus')).toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('Enter');
  await expect(menu).toBeHidden();
  await expect(more).toBeFocused();
  await expect(page.getByTestId('pptx-zoom')).toHaveValue('50%');

  await page.keyboard.press('ArrowDown');
  await expect(menu.locator('[data-label="Bold"]')).toHaveAttribute(
    'aria-disabled',
    'true'
  );
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(more).toBeFocused();
});

const LONGER = ' (übersetzt, mit deutlich längerer Beschriftung)';

type Strings = { [key: string]: string | Strings };

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
  const marker = /exportPng:[`"']Export PNG[`"']/;
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

const DEFAULT_TOOLBAR: [group: string, controls: string[]][] = [
  ['toolbar.groups.file', ['toolbar.save', 'toolbar.exportPng']],
  ['toolbar.groups.slides', ['toolbar.newSlide']],
  ['toolbar.groups.history', ['toolbar.undo', 'toolbar.redo']],
  ['toolbar.groups.zoom', ['toolbar.groups.zoom']],
  [
    'toolbar.groups.tools',
    [
      'commands.selectTool',
      'toolbar.textBoxTool',
      'toolbar.insertImage',
      'toolbar.shapeTool',
    ],
  ],
  [
    'toolbar.groups.font',
    [
      'toolbar.fontFamily',
      'toolbar.decreaseFontSize',
      'toolbar.fontSize',
      'toolbar.increaseFontSize',
    ],
  ],
  [
    'toolbar.groups.text',
    [
      'toolbar.bold',
      'toolbar.italic',
      'toolbar.underline',
      'toolbar.textColor',
    ],
  ],
  ['toolbar.groups.alignment', ['toolbar.groups.alignment']],
  [
    'toolbar.groups.shape',
    [
      'toolbar.fillColor',
      'toolbar.borderColor',
      'toolbar.borderWidth',
      'toolbar.arrange',
    ],
  ],
];

test('pptx: long translations keep every toolbar action reachable through More', async ({
  page,
}) => {
  const strings = JSON.parse(
    await readFile(resolve(root, 'packages/pptx-i18n/en.json'), 'utf8')
  ) as Strings;
  const long = (key: string) => lookup(strings, key) + LONGER;
  await lengthenStrings(page, strings);
  await page.setViewportSize({ width: 360, height: 800 });
  await open(page);

  const rail = page.getByTestId('pptx-formatting-toolbar');
  const more = rail.getByTestId('pptx-toolbar-more');
  await expect(more).toHaveAttribute('aria-label', long('toolbar.more'));
  await more.focus();
  await page.keyboard.press('ArrowDown');
  const menu = page.getByRole('menu', { name: long('toolbar.more') });
  await expect(menu).toBeVisible();

  const row = rail.locator('[data-toolbar-items]');
  let overflowed = 0;
  for (const [group, controls] of DEFAULT_TOOLBAR) {
    if (
      await row.getByRole('group', { name: long(group), exact: true }).count()
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
  expect(overflowed).toBeGreaterThan(15);
  expect(await fits(menu)).toBe(true);

  const arrange = menu.locator(`[data-label="${long('toolbar.arrange')}"]`);
  await expect(arrange).toHaveAttribute('aria-disabled', 'true');
  const reason = await arrange.evaluate(
    (item) =>
      document.getElementById(item.getAttribute('aria-describedby')!)
        ?.textContent
  );
  expect(reason).toBe(long('commands.reasons.objectRequired'));

  const zoom = menu.locator(`[data-label="${long('toolbar.groups.zoom')}"]`);
  await zoom.focus();
  await page.keyboard.press('ArrowRight');
  const levels = page.getByRole('menu', { name: long('toolbar.groups.zoom') });
  await expect(levels).toBeVisible();
  await expect(
    levels.locator(`[data-label="${long('toolbar.fit')}"]`)
  ).toHaveAttribute('aria-checked', 'true');
  expect(await fits(levels)).toBe(true);
});
