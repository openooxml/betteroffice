import { test, expect, type Locator, type Page } from 'playwright/test';
import JSZip from 'jszip';

const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const prefix = 'We collect feedback for ';
const followup = 'Measure it in week one.';
const input = (page: Page) => page.locator('.paged-editor__peerless-input');
const mirror = (page: Page) => page.locator('.canvas-page-mirror').first();
const canvas = (page: Page) => page.locator('.paged-editor__page canvas').first();
const cards = (page: Page) => page.locator('.docx-tracked-change-card.docx-review-card');
const dots = (page: Page) => page.locator('button.docx-review-dot');
const balloon = (page: Page) => page.locator('.docx-review-balloon');
const summary = (page: Page) => page.locator('.docx-review-summary');
const marks = (page: Page) => mirror(page).locator('.layout-run-text.docx-insertion, .layout-run-text.docx-deletion');

async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const body = `<w:p><w:r><w:t xml:space="preserve">${prefix}</w:t></w:r><w:del w:id="21" w:author="Agent" w:date="2026-01-01T10:00:00Z"><w:r><w:delText>six weeks</w:delText></w:r></w:del><w:ins w:id="22" w:author="Agent" w:date="2026-01-01T10:00:00Z"><w:r><w:t>four weeks</w:t></w:r></w:ins><w:r><w:t>, then ship.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Cut setup time in half.</w:t></w:r><w:ins w:id="23" w:author="Agent" w:date="2026-01-02T11:00:00Z"><w:r><w:t xml:space="preserve"> ${followup}</w:t></w:r></w:ins></w:p>
    <w:p><w:r><w:t xml:space="preserve">Merge edits </w:t></w:r><w:ins w:id="24" w:author="Lena" w:date="2026-01-03T12:00:00Z"><w:r><w:t xml:space="preserve">automatically </w:t></w:r></w:ins><w:r><w:t>when people reconnect.</w:t></w:r></w:p>`;
  zip.file('word/document.xml', `<w:document xmlns:w="${namespace}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, query = '/?format=docx') {
  await page.goto(query);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'review.docx', mimeType, buffer: await fixture(),
  });
  await expect(canvas(page)).toBeVisible({ timeout: 60_000 });
  await expect(input(page)).toBeEditable();
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(mirror(page)).toContainText('when people reconnect.');
  await expect(mirror(page)).toContainText(followup);
}

async function begin(page: Page, display: 'cards' | 'dots' = 'cards') {
  await open(page);
  await expect(cards(page).filter({ hasText: 'automatically' })).toBeVisible();
  await expect.poll(() => cards(page).count()).toBeGreaterThanOrEqual(3);
  const count = await cards(page).count();
  expect([3, 4]).toContain(count);
  if (display === 'dots') {
    await open(page, '/?format=docx&review=dots');
    await expect(dots(page)).toHaveCount(count);
  }
  return count;
}

async function mode(page: Page, target: 'Editing' | 'Suggesting') {
  const toolbar = page.getByTestId('formatting-bar');
  await toolbar.getByRole('button', { name: /^(Editing|Suggesting)$/, exact: true }).click();
  await page.getByRole('menuitemradio', { name: new RegExp(`^${target}`) }).click();
  await expect(toolbar.getByRole('button', { name: target, exact: true })).toBeVisible();
}

async function select(page: Page, start: number, length = 0) {
  await input(page).press('ControlOrMeta+Home');
  for (let index = 0; index < start; index++) await input(page).press('ArrowRight');
  for (let index = 0; index < length; index++) await input(page).press('Shift+ArrowRight');
}

async function showBalloon(page: Page, dot: Locator) {
  await dot.hover();
  await expect(balloon(page)).toBeVisible();
  await expect(dot).toHaveAttribute('aria-expanded', 'true');
  await expect(balloon(page).locator('.docx-review-card')).toHaveAttribute(
    'data-revision-id', (await dot.getAttribute('data-revision-id'))!,
  );
  return balloon(page);
}

async function expectSummary(page: Page, orientation: 'horizontal' | 'vertical', count: number) {
  await expect(summary(page)).toBeVisible();
  await expect(summary(page)).toHaveAttribute('data-orientation', orientation);
  await expect(summary(page).locator('.docx-review-summary__count')).toHaveText(new RegExp(`\\b${count}\\b`));
  const acceptAll = summary(page).locator('.docx-review-accept-all');
  await expect(acceptAll).toHaveAccessibleName('Accept all changes');
  await expect(acceptAll).toBeVisible();
}

test.beforeEach(async ({ page }) => {
  await page.setViewportSize({ width: 2560, height: 1200 });
});

test.describe('cards review', () => {
  test('keeps the page centred while the card column fits beside it and shifts it only as far as needed', async ({ page }) => {
    await begin(page);
    const geometry = () => page.evaluate(() => {
      const sheet = document.querySelector('.paged-editor__page')!.getBoundingClientRect();
      const aside = document.querySelector<HTMLElement>('.docx-unified-sidebar')!;
      const column = aside.getBoundingClientRect();
      const box = document.querySelector('.docx-editor__scroll-container')!.getBoundingClientRect();
      return { offset: Math.round(box.left + box.width / 2 - (sheet.left + sheet.width / 2)), room: Math.round(aside.offsetParent!.getBoundingClientRect().right - column.right) };
    });
    for (const width of [2560, 1920]) {
      await page.setViewportSize({ width, height: 1000 });
      await expect.poll(async () => Math.abs((await geometry()).offset)).toBeLessThanOrEqual(2);
      expect((await geometry()).room).toBeGreaterThanOrEqual(16);
    }
    await page.setViewportSize({ width: 1440, height: 900 });
    await expect.poll(async () => (await geometry()).room).toBeGreaterThanOrEqual(14);
    const { offset, room } = await geometry();
    expect(offset).toBeGreaterThan(0);
    expect(offset).toBeLessThan(176);
    expect(room).toBeLessThanOrEqual(18);
  });

  test('shows every change with decisions, human initials, and a horizontal summary', async ({ page }) => {
    const count = await begin(page);
    await expect(page.locator('.docx-review-dot')).toHaveCount(0);
    await expectSummary(page, 'horizontal', count);
    for (let index = 0; index < count; index++) {
      const card = cards(page).nth(index);
      await expect(card).toBeVisible();
      await expect(card).toHaveAttribute('data-kind', 'human');
      await expect(card).toHaveAttribute('data-active', /^(true|false)$/);
      await expect(card).toHaveAttribute('data-revision-id', /\S+/);
      await expect(card.locator('.docx-review-card__name')).toHaveText(index === count - 1 ? 'Lena' : 'Agent');
      await expect(card.locator('.docx-review-avatar')).toHaveText(index === count - 1 ? 'L' : 'A');
      await expect(card.locator('.docx-review-accept')).toHaveAccessibleName('Accept');
      await expect(card.locator('.docx-review-accept')).toBeVisible();
      await expect(card.locator('.docx-review-reject')).toHaveAccessibleName('Reject');
      await expect(card.locator('.docx-review-reject')).toBeVisible();
    }
  });

  test('activates cards and navigates in document order', async ({ page }) => {
    const count = await begin(page);
    const first = cards(page).first();
    const second = cards(page).nth(1);
    await first.locator('.docx-review-card__diff').click();
    await expect(first).toHaveAttribute('data-active', 'true');
    await expect(first.locator('.docx-review-nav__count')).toHaveText(`1/${count}`);
    await expect(page.locator('svg.docx-review-leader')).toBeVisible();
    await expect(first.locator('.docx-review-nav-next')).toHaveAccessibleName('Next change');
    await first.locator('.docx-review-nav-next').click();
    await expect(second).toHaveAttribute('data-active', 'true');
    await expect(first).toHaveAttribute('data-active', 'false');
    await expect(second.locator('.docx-review-nav__count')).toHaveText(`2/${count}`);
    await expect(second.locator('.docx-review-nav-prev')).toHaveAccessibleName('Previous change');
    await second.locator('.docx-review-nav-prev').click();
    await expect(first).toHaveAttribute('data-active', 'true');
    await expect(second).toHaveAttribute('data-active', 'false');
    await expect(first.locator('.docx-review-nav__count')).toHaveText(`1/${count}`);
  });

  test('styles replacement deletions red with a strike and insertions green', async ({ page }) => {
    const count = await begin(page);
    const replacement = cards(page).filter({ hasText: /six weeks|four weeks/ });
    await expect(replacement).toHaveCount(count - 2);
    const deleted = replacement.locator('.docx-review-card__diff .docx-review-diff-del');
    const inserted = replacement.locator('.docx-review-card__diff .docx-review-diff-ins');
    await expect(deleted).toHaveText('six weeks');
    await expect(inserted).toHaveText('four weeks');
    await expect(deleted).toHaveCSS('text-decoration-line', /line-through/);
    for (const [node, channel] of [[deleted, 'red'], [inserted, 'green']] as const) {
      await expect(node).toBeVisible();
      await expect.poll(() => node.evaluate((element, channel) => {
        const value = getComputedStyle(element).color;
        const channels = value.match(/^rgba?\(\s*([\d.]+)[,\s]+([\d.]+)/);
        if (!channels) throw new Error(`Expected an RGB color, received ${value}`);
        const red = Number(channels[1]);
        const green = Number(channels[2]);
        return channel === 'red' ? red - green : green - red;
      }, channel)).toBeGreaterThan(0);
    }
  });

  test('accepts Lena and rejects the independent follow-up insertion', async ({ page }) => {
    const count = await begin(page);
    await cards(page).filter({ hasText: 'automatically' }).locator('.docx-review-accept').click();
    await expect(cards(page)).toHaveCount(count - 1);
    await expect(mirror(page)).toContainText('Merge edits automatically when people reconnect.');
    await expect(mirror(page).locator('.docx-insertion').filter({ hasText: 'automatically' })).toHaveCount(0);
    await cards(page).filter({ hasText: followup }).locator('.docx-review-reject').click();
    await expect(cards(page)).toHaveCount(count - 2);
    await expect(mirror(page)).not.toContainText(followup);
    await expectSummary(page, 'horizontal', count - 2);
  });

  test('accepts all changes and hides the summary', async ({ page }) => {
    const count = await begin(page);
    await expectSummary(page, 'horizontal', count);
    await summary(page).locator('.docx-review-accept-all').click();
    await expect(cards(page)).toHaveCount(0);
    await expect(marks(page)).toHaveCount(0);
    await expect(summary(page)).toBeHidden();
    await expect(mirror(page)).toContainText(`${prefix}four weeks, then ship.`);
    await expect(mirror(page)).toContainText(followup);
    await expect(mirror(page)).toContainText('automatically');
    await expect(mirror(page)).not.toContainText('six weeks');
  });
});

test.describe('dots review', () => {
  test('keeps the page centred with one dot per change and a compact vertical summary', async ({ page }) => {
    const count = await begin(page, 'dots');
    await expect(page.locator('.docx-review-card:visible')).toHaveCount(0);
    await expect(page.locator('.docx-unified-sidebar .docx-review-card')).toHaveCount(0);
    await expect(page.locator('.docx-review-gutter')).toBeVisible();
    await expectSummary(page, 'vertical', count);
    await expect.poll(async () => {
      const bounds = await summary(page).locator('.docx-review-accept-all').boundingBox();
      return bounds?.width ?? Infinity;
    }).toBeLessThanOrEqual(40);
    await expect.poll(async () => {
      const surface = await canvas(page).boundingBox();
      const scroll = await page.locator('.docx-editor__scroll-container').boundingBox();
      return surface && scroll ? Math.abs(surface.x + surface.width / 2 - scroll.x - scroll.width / 2) : Infinity;
    }).toBeLessThanOrEqual(40);
    for (let index = 0; index < count; index++) {
      const dot = dots(page).nth(index);
      await expect(dot).toBeVisible();
      await expect(dot).toHaveAttribute('data-kind', 'human');
      await expect(dot).toHaveAttribute('data-active', /^(true|false)$/);
      await expect(dot).toHaveAttribute('data-revision-id', /\S+/);
      await expect(dot).toHaveAccessibleName(`Change by ${index === count - 1 ? 'Lena' : 'Agent'}`);
      await expect(dot).toHaveAttribute('aria-expanded', 'false');
    }
  });

  test('keeps a hovered balloon open through the pointer transition and closes on leave', async ({ page }) => {
    const count = await begin(page, 'dots');
    const first = dots(page).first();
    const panel = await showBalloon(page, first);
    await expect(panel.locator('.docx-review-card__diff')).toContainText('six weeks');
    if (count === 3) await expect(panel.locator('.docx-review-card__diff')).toContainText('four weeks');
    await expect(panel.locator('.docx-review-nav__count')).toHaveText(`1/${count}`);
    await expect(page.locator('svg.docx-review-leader')).toBeVisible();
    await panel.hover();
    await page.waitForTimeout(400);
    await expect(panel).toBeVisible();
    await expect(first).toHaveAttribute('aria-expanded', 'true');
    await page.mouse.move(1, 1);
    await expect(panel).toBeHidden({ timeout: 1000 });
    await expect(first).toHaveAttribute('aria-expanded', 'false');
  });

  test('opens on keyboard focus and restores focus to the dot on Escape', async ({ page }) => {
    const count = await begin(page, 'dots');
    const second = dots(page).nth(1);
    await second.focus();
    await expect(balloon(page)).toBeVisible();
    await expect(second).toHaveAttribute('aria-expanded', 'true');
    await expect(balloon(page).locator('.docx-review-nav__count')).toHaveText(`2/${count}`);
    await expect(balloon(page).locator('.docx-review-card')).toHaveAttribute(
      'data-revision-id', (await second.getAttribute('data-revision-id'))!,
    );
    await page.keyboard.press('Escape');
    await expect(balloon(page)).toBeHidden();
    await expect(second).toHaveAttribute('aria-expanded', 'false');
    await expect(second).toBeFocused();
  });

  test('accepts Lena and rejects the follow-up through their balloons', async ({ page }) => {
    const count = await begin(page, 'dots');
    const lena = dots(page).and(page.getByRole('button', { name: 'Change by Lena', exact: true }));
    await (await showBalloon(page, lena)).locator('.docx-review-accept').click();
    await expect(dots(page)).toHaveCount(count - 1);
    await expect(lena).toHaveCount(0);
    await expect(mirror(page)).toContainText('Merge edits automatically when people reconnect.');
    await expect(mirror(page).locator('.docx-insertion').filter({ hasText: 'automatically' })).toHaveCount(0);
    const followupDot = dots(page).last();
    const panel = await showBalloon(page, followupDot);
    await expect(panel.locator('.docx-review-card__diff')).toContainText(followup);
    await panel.locator('.docx-review-reject').click();
    await expect(dots(page)).toHaveCount(count - 2);
    await expect(mirror(page)).not.toContainText(followup);
    await expectSummary(page, 'vertical', count - 2);
  });

  test('navigates from the first balloon to the next change', async ({ page }) => {
    const count = await begin(page, 'dots');
    const first = dots(page).first();
    const second = dots(page).nth(1);
    const panel = await showBalloon(page, first);
    await expect(panel.locator('.docx-review-nav__count')).toHaveText(`1/${count}`);
    await panel.locator('.docx-review-nav-next').click();
    await expect(panel).toBeVisible();
    await expect(panel.locator('.docx-review-nav__count')).toHaveText(`2/${count}`);
    await expect(panel.locator('.docx-review-card')).toHaveAttribute(
      'data-revision-id', (await second.getAttribute('data-revision-id'))!,
    );
    await expect(second).toHaveAttribute('data-active', 'true');
    await expect(first).toHaveAttribute('data-active', 'false');
  });

  test('opens the replacement balloon when the caret selects its inserted text', async ({ page }) => {
    const count = await begin(page, 'dots');
    await mode(page, 'Editing');
    await canvas(page).click({ position: { x: 150, y: 150 } });
    await page.mouse.move(1, 1);
    await select(page, prefix.length + 'six weeks'.length + 2, 2);
    await expect(balloon(page)).toBeVisible();
    await expect(balloon(page).locator('.docx-review-card__diff')).toContainText('four weeks');
    const index = count === 3 ? 0 : 1;
    await expect(balloon(page).locator('.docx-review-nav__count')).toHaveText(`${index + 1}/${count}`);
    await expect(dots(page).nth(index)).toHaveAttribute('data-active', 'true');
    await expect(dots(page).nth(index)).toHaveAttribute('aria-expanded', 'true');
  });

  test('accepts all changes from the vertical summary', async ({ page }) => {
    const count = await begin(page, 'dots');
    await expectSummary(page, 'vertical', count);
    await summary(page).locator('.docx-review-accept-all').click();
    await expect(dots(page)).toHaveCount(0);
    await expect(marks(page)).toHaveCount(0);
    await expect(balloon(page)).toBeHidden();
    await expect(summary(page)).toBeHidden();
    await expect(mirror(page)).toContainText(`${prefix}four weeks, then ship.`);
    await expect(mirror(page)).toContainText(followup);
    await expect(mirror(page)).toContainText('automatically');
    await expect(mirror(page)).not.toContainText('six weeks');
  });
});
