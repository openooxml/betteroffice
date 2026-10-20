import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const original = 'Alpha beta gamma';
const insertion = 'Suggested ';
const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const input = (page: Page) => page.locator('.paged-editor__peerless-input');
const mirror = (page: Page) => page.locator('.canvas-page-mirror').first();
const cards = (page: Page) => page.locator('.docx-tracked-change-card');
const toolbar = (page: Page) => page.getByTestId('formatting-bar');

async function fixture(review = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const body = review
    ? '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:ins w:id="11" w:author="Writer" w:date="2026-01-01T10:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins></w:p><w:p><w:r><w:t xml:space="preserve">Beta </w:t></w:r><w:del w:id="12" w:author="Reviewer" w:date="2026-01-02T11:00:00Z"><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>'
    : `<w:p><w:r><w:t>${original}</w:t></w:r></w:p>`;
  zip.file('word/document.xml', `<w:document xmlns:w="${namespace}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, source: Buffer | string) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles(typeof source === 'string'
    ? source
    : { name: 'tracked.docx', mimeType, buffer: source });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(input(page)).toBeEditable();
  await expect(page.locator('.notice')).toHaveCount(0);
}

async function mode(page: Page, target: 'Editing' | 'Suggesting') {
  await toolbar(page).getByRole('button', { name: /^(Editing|Suggesting)$/, exact: true }).click();
  await page.getByRole('menuitemradio', { name: new RegExp(`^${target}`) }).click();
  await expect(toolbar(page).getByRole('button', { name: target, exact: true })).toBeVisible();
}

async function select(page: Page, start: number, length = 0) {
  await input(page).press('ControlOrMeta+Home');
  for (let index = 0; index < start; index++) await input(page).press('ArrowRight');
  for (let index = 0; index < length; index++) await input(page).press('Shift+ArrowRight');
}

async function begin(page: Page) {
  await open(page, await fixture());
  await expect(mirror(page)).toContainText(original);
  await mode(page, 'Suggesting');
  await page.locator('.paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
  await select(page, 0);
}

async function revision(page: Page, kind: 'ins' | 'del', text: string) {
  const marks = mirror(page).locator(kind === 'ins' ? '.layout-run-text.docx-insertion' : '.layout-run-text.docx-deletion');
  await expect.poll(async () => (await marks.allTextContents()).join('')).toBe(text);
  await expect(marks.first()).toHaveAttribute('role', kind === 'ins' ? 'insertion' : 'deletion');
  await expect(marks.first()).toHaveAttribute('data-revision-id', /\S+/);
  await expect(marks.first()).toHaveAttribute('data-change-author', /\S+/);
  await expect(marks.first()).toHaveAttribute('data-change-date', /^\d{4}-\d{2}-\d{2}T/);
  await expect(marks.first()).toHaveAttribute('aria-description', /\S+/);
  await expect(cards(page).filter({ hasText: text.trim() }).first()).toBeVisible();
}

async function save(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  const state = await page.evaluate(({ xml, namespace }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid document XML');
    const nodes = (name: string, parent: Document | Element = document) => [...parent.getElementsByTagNameNS(namespace, name)];
    const within = (node: Element, names: string[]): boolean => {
      for (let parent = node.parentElement; parent; parent = parent.parentElement) {
        if (parent.namespaceURI === namespace && names.includes(parent.localName)) return true;
      }
      return false;
    };
    const changes = (name: string) => nodes(name).map(node => ({
      id: node.getAttributeNS(namespace, 'id'),
      author: node.getAttributeNS(namespace, 'author'),
      date: node.getAttributeNS(namespace, 'date'),
      text: [...nodes('t', node), ...nodes('delText', node)].map(node => node.textContent ?? '').join(''),
      deletedRuns: nodes('delText', node).length,
      normalRuns: nodes('t', node).length,
    }));
    const text = nodes('t').filter(node => !within(node, ['del', 'rPrChange']));
    return {
      insertions: changes('ins'),
      deletions: changes('del'),
      formatting: changes('rPrChange'),
      text: text.map(node => node.textContent ?? '').join(''),
      plain: text.filter(node => !within(node, ['ins'])).map(node => node.textContent ?? '').join(''),
      deleted: nodes('delText').map(node => node.textContent ?? '').join(''),
      bold: nodes('r').filter(node => {
        const properties = [...node.children].find(child => child.localName === 'rPr');
        return properties && [...properties.children].some(child => child.localName === 'b' &&
          !['0', 'false', 'off'].includes(child.getAttributeNS(namespace, 'val') ?? '1'));
      }).map(node => nodes('t', node).map(node => node.textContent ?? '').join('')).join(''),
    };
  }, { xml, namespace });
  const changes = [...state.insertions, ...state.deletions, ...state.formatting];
  for (const change of changes) {
    expect(change.id).toMatch(/^\d+$/);
    expect(change.author?.trim()).toBeTruthy();
    expect(change.date).toMatch(/^\d{4}-\d{2}-\d{2}T/);
    expect(Number.isNaN(Date.parse(change.date!))).toBe(false);
  }
  for (const deletion of state.deletions) {
    expect(deletion.deletedRuns).toBeGreaterThan(0);
    expect(deletion.normalRuns).toBe(0);
  }
  return { file, state };
}

type Saved = Awaited<ReturnType<typeof save>>;

async function reopen(page: Page, info: TestInfo, saved: Saved) {
  await open(page, saved.file);
  await expect(mirror(page)).toContainText(saved.state.text.slice(0, 3));
  const inserted = saved.state.insertions.map(change => change.text).join('');
  if (inserted) await revision(page, 'ins', inserted);
  if (saved.state.deleted) await revision(page, 'del', saved.state.deleted);
  if (!inserted && !saved.state.deleted && !saved.state.formatting.length) {
    await expect(cards(page)).toHaveCount(0);
    await expect(mirror(page).locator('.docx-insertion, .docx-deletion')).toHaveCount(0);
  }
  expect((await save(page, info, 'reopened')).state).toEqual(saved.state);
}

async function decide(page: Page, text: string, decision: 'Accept' | 'Reject') {
  const card = cards(page).filter({ hasText: text });
  await expect(card).toBeVisible();
  await card.click();
  await card.getByTitle(decision, { exact: true }).click();
  await expect(card).toHaveCount(0);
}

async function undo(page: Page) {
  await input(page).press('ControlOrMeta+z');
}

async function redo(page: Page) {
  await input(page).press('ControlOrMeta+Shift+z');
}

test.beforeEach(async ({ page }) => {
  await page.setViewportSize({ width: 2560, height: 1200 });
});

test('suggesting records an insertion, supports undo and redo, and reopens its revision', async ({ page }, info) => {
  await begin(page);
  await page.keyboard.insertText(insertion);
  await revision(page, 'ins', insertion);
  await expect(cards(page)).toHaveCount(1);
  const edited = await save(page, info, 'insertion');
  expect(edited.state.insertions.map(change => change.text).join('')).toBe(insertion);
  expect(edited.state.text).toBe(insertion + original);
  expect(edited.state.plain).toBe(original);
  expect(edited.state.deletions).toEqual([]);
  await undo(page);
  await expect(cards(page)).toHaveCount(0);
  await expect(mirror(page)).not.toContainText(insertion);
  const undone = await save(page, info, 'insertion-undone');
  expect(undone.state.text).toBe(original);
  expect(undone.state.insertions).toEqual([]);
  await redo(page);
  await revision(page, 'ins', insertion);
  expect((await save(page, info, 'insertion-redone')).state).toEqual(edited.state);
  await reopen(page, info, edited);
});

test('deleting a selected word preserves red deleted text and w:delText through reopen', async ({ page }, info) => {
  await begin(page);
  await select(page, 6, 4);
  await input(page).press('Delete');
  await revision(page, 'del', 'beta');
  const edited = await save(page, info, 'selection-deleted');
  expect(edited.state.text).toBe('Alpha  gamma');
  expect(edited.state.deleted).toBe('beta');
  expect(edited.state.insertions).toEqual([]);
  await undo(page);
  await expect(cards(page)).toHaveCount(0);
  await expect(mirror(page).locator('.docx-deletion')).toHaveCount(0);
  expect((await save(page, info, 'selection-undone')).state.text).toBe(original);
  await redo(page);
  await revision(page, 'del', 'beta');
  expect((await save(page, info, 'selection-redone')).state).toEqual(edited.state);
  await reopen(page, info, edited);
});

test('Backspace across a word boundary tracks every deleted character and supports undo and redo', async ({ page }, info) => {
  await begin(page);
  await select(page, 10);
  for (let index = 0; index < 7; index++) await input(page).press('Backspace');
  await revision(page, 'del', 'ha beta');
  const edited = await save(page, info, 'word-deleted');
  expect(edited.state.text).toBe('Alp gamma');
  expect(edited.state.deleted).toBe('ha beta');
  expect(edited.state.insertions).toEqual([]);
  let undone = edited;
  const deletedStates = [edited.state.deleted];
  while (undone.state.deleted) {
    expect(deletedStates.length).toBeLessThanOrEqual(edited.state.deleted.length);
    const previousDeleted = undone.state.deleted;
    await undo(page);
    await expect.poll(async () => (await mirror(page).locator('.docx-deletion').allTextContents()).join('')).not.toBe(previousDeleted);
    undone = await save(page, info, `word-undone-${deletedStates.length}`);
    expect(undone.state.deleted.length).toBeLessThan(previousDeleted.length);
    expect(undone.state.deleted).toBe(previousDeleted.slice(previousDeleted.length - undone.state.deleted.length));
    expect(undone.state.text).toBe(original.slice(0, 10 - undone.state.deleted.length) + original.slice(10));
    deletedStates.push(undone.state.deleted);
  }
  expect(undone.state.deleted).toBe('');
  expect(undone.state.text).toBe(original);
  expect(undone.state.deletions).toEqual([]);
  await expect(cards(page)).toHaveCount(0);
  for (let index = deletedStates.length - 2; index >= 0; index--) {
    await redo(page);
    await expect.poll(async () => (await mirror(page).locator('.docx-deletion').allTextContents()).join('')).toBe(deletedStates[index]);
  }
  await revision(page, 'del', 'ha beta');
  expect((await save(page, info, 'word-redone')).state).toEqual(edited.state);
  await reopen(page, info, edited);
});

for (const decision of ['Accept', 'Reject'] as const) {
  test(`${decision.toLowerCase()} a replacement groups both revisions and supports undo, redo, and reopen`, async ({ page }, info) => {
    await begin(page);
    await select(page, 6, 4);
    await page.keyboard.insertText('Delta');
    await revision(page, 'ins', 'Delta');
    await revision(page, 'del', 'beta');
    await expect(cards(page)).toHaveCount(1);
    await expect(cards(page)).toContainText('Replaced');
    await expect(cards(page)).toContainText('beta');
    await expect(cards(page)).toContainText('Delta');
    const edited = await save(page, info, 'replacement');
    expect(edited.state.text).toBe('Alpha Delta gamma');
    expect(edited.state.deleted).toBe('beta');
    await undo(page);
    await expect(cards(page)).toHaveCount(0);
    const undone = await save(page, info, 'replacement-undone');
    expect(undone.state.text).toBe(original);
    expect(undone.state.insertions).toEqual([]);
    expect(undone.state.deletions).toEqual([]);
    await redo(page);
    await revision(page, 'ins', 'Delta');
    await revision(page, 'del', 'beta');
    expect((await save(page, info, 'replacement-redone')).state).toEqual(edited.state);
    await reopen(page, info, edited);
    await expect(cards(page)).toHaveCount(1);
    await decide(page, 'Delta', decision);
    await expect(cards(page)).toHaveCount(0);
    await expect(mirror(page).locator('.docx-insertion, .docx-deletion')).toHaveCount(0);
    if (decision === 'Accept') await expect(mirror(page)).toContainText('Delta');
    else await expect(mirror(page)).toContainText(original);
    const resolved = await save(page, info, 'replacement-resolved');
    expect(resolved.state.insertions).toEqual([]);
    expect(resolved.state.deletions).toEqual([]);
    expect(resolved.state.text).toBe(decision === 'Accept' ? 'Alpha Delta gamma' : original);
    await undo(page);
    await revision(page, 'ins', 'Delta');
    await revision(page, 'del', 'beta');
    expect((await save(page, info, 'replacement-decision-undone')).state).toEqual(edited.state);
    await redo(page);
    await expect(cards(page)).toHaveCount(0);
    expect((await save(page, info, 'replacement-decision-redone')).state).toEqual(resolved.state);
    await reopen(page, info, resolved);
  });
}

for (const decision of ['Accept', 'Reject'] as const) {
  test(`${decision.toLowerCase()} an insertion from its card, undo and redo the decision, and reopen`, async ({ page }, info) => {
    await begin(page);
    await page.keyboard.insertText(insertion);
    await revision(page, 'ins', insertion);
    const pending = await save(page, info, 'insertion-pending');
    await decide(page, insertion.trim(), decision);
    await expect(mirror(page).locator('.docx-insertion')).toHaveCount(0);
    if (decision === 'Accept') await expect(mirror(page)).toContainText(insertion);
    else await expect(mirror(page)).not.toContainText(insertion);
    const resolved = await save(page, info, 'insertion-resolved');
    expect(resolved.state.insertions).toEqual([]);
    expect(resolved.state.deletions).toEqual([]);
    expect(resolved.state.text).toBe(decision === 'Accept' ? insertion + original : original);
    expect(resolved.state.plain).toBe(resolved.state.text);
    await undo(page);
    await revision(page, 'ins', insertion);
    expect((await save(page, info, 'decision-undone')).state).toEqual(pending.state);
    await redo(page);
    await expect(cards(page)).toHaveCount(0);
    expect((await save(page, info, 'decision-redone')).state).toEqual(resolved.state);
    await reopen(page, info, resolved);
  });
}

test('accepting one revision and rejecting another leaves their independent decisions saved', async ({ page }, info) => {
  await open(page, await fixture(true));
  await revision(page, 'ins', 'added');
  await revision(page, 'del', 'removed');
  await expect(cards(page)).toHaveCount(2);
  await decide(page, 'added', 'Accept');
  await expect(cards(page)).toHaveCount(1);
  await revision(page, 'del', 'removed');
  const accepted = await save(page, info, 'one-accepted');
  expect(accepted.state.insertions).toEqual([]);
  expect(accepted.state.deletions).toHaveLength(1);
  expect(accepted.state.text).toBe('Alpha addedBeta ');
  await decide(page, 'removed', 'Reject');
  await expect(cards(page)).toHaveCount(0);
  await expect(mirror(page)).toContainText('added');
  await expect(mirror(page)).toContainText('removed');
  const resolved = await save(page, info, 'mixed-decisions');
  expect(resolved.state.text).toBe('Alpha addedBeta removed');
  expect(resolved.state.insertions).toEqual([]);
  expect(resolved.state.deletions).toEqual([]);
  await undo(page);
  await revision(page, 'del', 'removed');
  expect((await save(page, info, 'reject-undone')).state).toEqual(accepted.state);
  await redo(page);
  await expect(cards(page)).toHaveCount(0);
  expect((await save(page, info, 'reject-redone')).state).toEqual(resolved.state);
  await reopen(page, info, resolved);
});

test('turning suggesting off leaves later typing untracked while retaining the earlier revision', async ({ page }, info) => {
  await begin(page);
  await page.keyboard.insertText(insertion);
  await revision(page, 'ins', insertion);
  await mode(page, 'Editing');
  await input(page).press('ControlOrMeta+End');
  await page.keyboard.insertText(' plain');
  await expect(mirror(page)).toContainText('gamma plain');
  await revision(page, 'ins', insertion);
  await expect(cards(page)).toHaveCount(1);
  const edited = await save(page, info, 'tracking-off');
  expect(edited.state.text).toBe(insertion + original + ' plain');
  expect(edited.state.plain).toBe(original + ' plain');
  expect(edited.state.insertions.map(change => change.text).join('')).toBe(insertion);
  await reopen(page, info, edited);
});
