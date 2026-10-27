import { test, expect, type Locator, type Page } from 'playwright/test';
import JSZip from 'jszip';

const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const prefix = 'We collect feedback for ';
const followup = 'Measure it in week one.';
const input = (page: Page) => page.locator('.paged-editor__input');
const mirror = (page: Page) => page.locator('.canvas-page-mirror').first();
const canvas = (page: Page) => page.locator('.paged-editor__page canvas').first();
const cards = (page: Page) => page.locator('.docx-tracked-change-card.docx-review-card');
const dots = (page: Page) => page.locator('button.docx-review-dot');
const balloon = (page: Page) => page.locator('.docx-review-balloon');
const summary = (page: Page) => page.locator('.docx-review-summary');
const marks = (page: Page) => mirror(page).locator('.layout-run-text.docx-insertion, .layout-run-text.docx-deletion');

const spaced = '<w:pPr><w:spacing w:after="1200"/></w:pPr>';
const defaultBody = `<w:p>${spaced}<w:r><w:t xml:space="preserve">${prefix}</w:t></w:r><w:del w:id="21" w:author="Agent" w:date="2026-01-01T10:00:00Z"><w:r><w:delText>six weeks</w:delText></w:r></w:del><w:ins w:id="22" w:author="Agent" w:date="2026-01-01T10:00:00Z"><w:r><w:t>four weeks</w:t></w:r></w:ins><w:r><w:t>, then ship.</w:t></w:r></w:p>
    <w:p>${spaced}<w:r><w:t>Cut setup time in half.</w:t></w:r><w:ins w:id="23" w:author="Agent" w:date="2026-01-02T11:00:00Z"><w:r><w:t xml:space="preserve"> ${followup}</w:t></w:r></w:ins></w:p>
    <w:p>${spaced}<w:r><w:t xml:space="preserve">Merge edits </w:t></w:r><w:ins w:id="24" w:author="Lena" w:date="2026-01-03T12:00:00Z"><w:r><w:t xml:space="preserve">automatically </w:t></w:r></w:ins><w:r><w:t>when people reconnect.</w:t></w:r></w:p>`;
const longName = 'Alexandra Konstantinopoulou-Bergström';
const far = '<w:pPr><w:spacing w:after="2400"/></w:pPr>';
const sameLineBody = `<w:p>${far}<w:r><w:t xml:space="preserve">Owners: </w:t></w:r><w:ins w:id="31" w:author="Lena" w:date="2026-01-03T12:00:00Z"><w:r><w:t>Mia</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:del w:id="32" w:author="Agent" w:date="2026-01-03T12:00:00Z"><w:r><w:delText>Ravi</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> review the plan.</w:t></w:r></w:p>
    <w:p>${spaced}<w:r><w:t xml:space="preserve">Merge edits </w:t></w:r><w:ins w:id="33" w:author="${longName}" w:date="2026-01-03T12:00:00Z"><w:r><w:t xml:space="preserve">automatically </w:t></w:r></w:ins><w:r><w:t>when people reconnect.</w:t></w:r></w:p>`;

async function fixture(body = defaultBody) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="${namespace}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, query = '/?format=docx', body = defaultBody) {
  await page.goto(query);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'review.docx', mimeType, buffer: await fixture(body),
  });
  await expect(canvas(page)).toBeVisible({ timeout: 60_000 });
  await expect(input(page)).toBeEditable();
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(mirror(page)).toContainText('when people reconnect.');
  if (body === defaultBody) await expect(mirror(page)).toContainText(followup);
}

async function begin(page: Page, display: 'cards' | 'dots' = 'cards') {
  await open(page);
  await expect(cards(page).filter({ hasText: 'Lena' })).toBeVisible();
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

async function expectMinimalCard(card: Locator) {
  await expect(card.locator('.docx-review-card__diff, del, ins')).toHaveCount(0);
  for (const [selector, name] of [['.docx-review-accept', 'Accept change'], ['.docx-review-reject', 'Reject change']] as const) {
    const button = card.locator(`.docx-review-card__header .docx-review-card__actions ${selector}`);
    await expect(button).toBeVisible();
    await expect(button).toHaveAccessibleName(name);
    await expect(button).toHaveAttribute('title', name);
    await expect(button).toHaveText('');
    await expect(button.locator('svg[viewBox="0 0 24 24"] path').first()).toBeAttached();
  }
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
    expect(Math.abs((await geometry()).offset)).toBeLessThanOrEqual(16);
    await page.setViewportSize({ width: 1280, height: 900 });
    await expect.poll(async () => (await geometry()).room).toBeGreaterThanOrEqual(14);
    const { offset, room } = await geometry();
    expect(offset).toBeGreaterThan(0);
    expect(offset).toBeLessThan(126);
    expect(room).toBeLessThanOrEqual(18);
  });

  test('keeps cards and the review-only column within 240px', async ({ page }) => {
    await begin(page);
    expect((await page.locator('.docx-unified-sidebar').boundingBox())!.width).toBeLessThanOrEqual(240);
    for (const bounds of await cards(page).evaluateAll(nodes => nodes.map(node => node.getBoundingClientRect().width))) expect(bounds).toBeLessThanOrEqual(240);
    const lena = cards(page).filter({ hasText: 'Lena' });
    const header = (await lena.locator('.docx-review-card__header').boundingBox())!;
    const actions = (await lena.locator('.docx-review-card__actions').boundingBox())!;
    const meta = (await lena.locator('.docx-review-card__meta').boundingBox())!;
    expect(actions.x - (meta.x + meta.width)).toBeLessThanOrEqual(12);
    expect(header.x + header.width - (actions.x + actions.width)).toBeLessThanOrEqual(10);
  });

  test('draws the dotted leader for the hovered and the active card, never for every card', async ({ page }) => {
    await begin(page);
    const leaders = page.locator('svg.docx-review-leader');
    await page.mouse.move(1, 1);
    const idle = await leaders.count();
    expect(idle).toBeLessThanOrEqual(1);
    const lena = cards(page).filter({ hasText: 'Lena' });
    await lena.locator('.docx-review-card__name').hover();
    await expect(leaders).toHaveCount(idle + 1);
    await expect(leaders.last().locator('path')).toHaveCSS('stroke-dasharray', /\d/);
    await page.mouse.move(1, 1);
    await expect(leaders).toHaveCount(idle);
    await cards(page).first().locator('.docx-review-card__name').click();
    await expect(leaders).toHaveCount(1);
    await lena.locator('.docx-review-card__name').hover();
    await expect(leaders).toHaveCount(2);
  });

  test('shows the change in a lean row only for cards that share a line, without overlapping cards', async ({ page }) => {
    await open(page, '/?format=docx', sameLineBody);
    await expect(cards(page)).toHaveCount(3);
    const rows = cards(page).locator('.docx-review-card__diff');
    await expect(rows).toHaveCount(2);
    for (const index of [0, 1]) {
      await expect(cards(page).nth(index)).toHaveAttribute('data-ambiguous', 'true');
      const row = (await rows.nth(index).boundingBox())!, header = (await cards(page).nth(index).locator('.docx-review-card__header').boundingBox())!;
      expect(row.height).toBeLessThanOrEqual(26);
      expect(row.y).toBeGreaterThanOrEqual(header.y + header.height - 1);
    }
    await expect(rows.nth(0).locator('ins.docx-review-diff-ins')).toHaveText('Mia');
    await expect(rows.nth(0).locator('ins.docx-review-diff-ins')).toHaveCSS('text-decoration-style', 'dotted');
    await expect(rows.nth(1).locator('del.docx-review-diff-del')).toHaveText('Ravi');
    await expect(rows.nth(1).locator('del.docx-review-diff-del')).toHaveCSS('text-decoration-line', 'line-through');
    await expect(cards(page).nth(2)).toHaveAttribute('data-ambiguous', 'false');
    await expectMinimalCard(cards(page).nth(2));
    const name = cards(page).nth(2).locator('.docx-review-card__name');
    await expect(name).toHaveText(longName);
    await expect(name).toHaveCSS('text-overflow', 'ellipsis');
    expect(await name.evaluate(node => node.scrollWidth > node.clientWidth)).toBe(true);
    await expect.poll(async () => {
      const boxes = await cards(page).evaluateAll(nodes => nodes.map(node => node.getBoundingClientRect().toJSON() as DOMRect));
      return boxes.every(box => box.width <= 240) && boxes.every((box, index) => index === 0 || box.top >= boxes[index - 1]!.bottom);
    }).toBe(true);
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
      if (count === 3) await expectMinimalCard(card);
    }
  });

  test('activates cards and navigates in document order', async ({ page }) => {
    const count = await begin(page);
    const first = cards(page).first();
    const second = cards(page).nth(1);
    await first.locator('.docx-review-card__name').click();
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

  test('keeps each card to one row with small icon-only accept and reject that hover green and red and show a focus ring', async ({ page }) => {
    await begin(page);
    const card = cards(page).filter({ hasText: 'Lena' });
    const header = await card.locator('.docx-review-card__header').boundingBox();
    expect((await card.boundingBox())!.height).toBeLessThanOrEqual(header!.height + 2);
    for (const [selector, channel] of [['.docx-review-accept', 'green'], ['.docx-review-reject', 'red']] as const) {
      const button = card.locator(selector);
      const bounds = (await button.boundingBox())!;
      expect(bounds.width).toBeLessThanOrEqual(24);
      expect(bounds.height).toBeLessThanOrEqual(24);
      expect(Math.abs(bounds.y + bounds.height / 2 - header!.y - header!.height / 2)).toBeLessThanOrEqual(2);
      await button.hover();
      await expect.poll(() => button.evaluate((element, channel) => {
        const value = getComputedStyle(element).color;
        const channels = value.match(/^rgba?\(\s*([\d.]+)[,\s]+([\d.]+)/);
        if (!channels) throw new Error(`Expected an RGB color, received ${value}`);
        return channel === 'red' ? Number(channels[1]) - Number(channels[2]) : Number(channels[2]) - Number(channels[1]);
      }, channel)).toBeGreaterThan(0);
    }
    await page.mouse.move(1, 1);
    await card.locator('.docx-review-accept').focus();
    await page.keyboard.press('Tab');
    await expect(card.locator('.docx-review-reject')).toBeFocused();
    await expect(card.locator('.docx-review-reject')).toHaveCSS('outline-style', 'solid');
  });

  test('accepts Lena and rejects the independent follow-up insertion', async ({ page }) => {
    const count = await begin(page);
    await cards(page).filter({ hasText: 'Lena' }).getByRole('button', { name: 'Accept change' }).click();
    await expect(cards(page)).toHaveCount(count - 1);
    await expect(mirror(page)).toContainText('Merge edits automatically when people reconnect.');
    await expect(mirror(page).locator('.docx-insertion').filter({ hasText: 'automatically' })).toHaveCount(0);
    await cards(page).last().getByRole('button', { name: 'Reject change' }).click();
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
    await expectMinimalCard(panel.locator('.docx-review-card'));
    await expect(panel.locator('.docx-review-card__name')).toHaveText('Agent');
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

  for (const width of [1440, 1280]) test(`places the balloon so it never covers the anchor text at ${width}`, async ({ page }) => {
    await page.setViewportSize({ width, height: 900 });
    const count = await begin(page, 'dots');
    for (const [index, text] of [[0, 'six weeks'], [count - 1, 'automatically']] as const) {
      await page.mouse.move(1, 1);
      await expect(balloon(page)).toBeHidden();
      const dot = dots(page).nth(index);
      const panel = (await (await showBalloon(page, dot)).boundingBox())!;
      const anchor = (await dot.boundingBox())!;
      const words = await marks(page).evaluateAll((nodes, target) => {
        const joined = nodes.map(node => node.textContent ?? '').join('');
        let offset = 0;
        const start = joined.indexOf(target), end = start + target.length;
        return nodes.flatMap(node => {
          const from = offset;
          offset += (node.textContent ?? '').length;
          return start >= 0 && from < end && offset > start && node.textContent!.trim() ? [node.getBoundingClientRect().toJSON() as DOMRect] : [];
        });
      }, text);
      expect(words.length).toBeGreaterThan(0);
      const below = panel.y >= anchor.y + anchor.height;
      for (const word of words) {
        expect(word.width).toBeGreaterThan(0);
        expect(below ? panel.y >= word.bottom : panel.x >= word.right).toBe(true);
      }
      if (below) {
        expect(panel.x).toBeLessThan(anchor.x);
        expect(panel.x + panel.width).toBeGreaterThanOrEqual(anchor.x + anchor.width - 1);
      } else expect(panel.x).toBeGreaterThanOrEqual(anchor.x + anchor.width);
      if (width === 1280) expect(below).toBe(true);
      expect(panel.x + panel.width).toBeLessThanOrEqual(width);
    }
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
    await (await showBalloon(page, lena)).getByRole('button', { name: 'Accept change' }).click();
    await expect(dots(page)).toHaveCount(count - 1);
    await expect(lena).toHaveCount(0);
    await expect(mirror(page)).toContainText('Merge edits automatically when people reconnect.');
    await expect(mirror(page).locator('.docx-insertion').filter({ hasText: 'automatically' })).toHaveCount(0);
    const followupDot = dots(page).last();
    const panel = await showBalloon(page, followupDot);
    await expect(panel.locator('.docx-review-card__name')).toHaveText('Agent');
    await panel.getByRole('button', { name: 'Reject change' }).click();
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
    const index = count === 3 ? 0 : 1;
    await expect(balloon(page).locator('.docx-review-card')).toHaveAttribute(
      'data-revision-id', (await dots(page).nth(index).getAttribute('data-revision-id'))!,
    );
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

test.describe('in-thread route', () => {
  async function openInThread(page: Page, query = '') {
    await page.goto(`/?format=docx&peerless=0${query}`);
    await page.locator('input[type=file]').last().setInputFiles({ name: 'review.docx', mimeType, buffer: await fixture() });
    await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  }

  test('cards show the summary pill and the dotted leader for the active card', async ({ page }) => {
    await openInThread(page);
    const lena = cards(page).filter({ hasText: 'Lena' });
    await expect(lena).toBeVisible({ timeout: 30_000 });
    await expect(page.locator('.docx-review-dot')).toHaveCount(0);
    await expectSummary(page, 'horizontal', await cards(page).count());
    await lena.locator('.docx-review-card__name').click();
    await expect(lena).toHaveAttribute('data-active', 'true');
    await expect(page.locator('svg.docx-review-leader path')).toHaveAttribute('d', /^M[\d.-]+ [\d.-]+ H/);
    const card = (await lena.boundingBox())!;
    const end = await page.locator('svg.docx-review-leader path').evaluate(node => {
      const path = node as unknown as SVGPathElement;
      const points = path.getAttribute('d')!.match(/-?[\d.]+/g)!.map(Number);
      const box = path.ownerSVGElement!.getBoundingClientRect();
      return { x: box.left + points.at(-2)!, y: box.top + points.at(-1)! };
    });
    expect(Math.abs(end.x - card.x)).toBeLessThanOrEqual(2);
    expect(end.y).toBeGreaterThanOrEqual(card.y);
    expect(end.y).toBeLessThanOrEqual(card.y + card.height);
  });

  test('dots mode shows a dot per change, a vertical summary and accepts from the balloon', async ({ page }) => {
    await openInThread(page, '&review=dots');
    await expect.poll(() => dots(page).count(), { timeout: 30_000 }).toBeGreaterThanOrEqual(3);
    const count = await dots(page).count();
    await expect(page.locator('.docx-unified-sidebar .docx-review-card')).toHaveCount(0);
    await expectSummary(page, 'vertical', count);
    const lena = dots(page).last();
    await expect(lena).toHaveAccessibleName('Change by Lena');
    const panel = await showBalloon(page, lena);
    await expect(page.locator('svg.docx-review-leader')).toBeVisible();
    await panel.getByRole('button', { name: 'Accept change' }).click();
    await expect(dots(page)).toHaveCount(count - 1);
  });
});
