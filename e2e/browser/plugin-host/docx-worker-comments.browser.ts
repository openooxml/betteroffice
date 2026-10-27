import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { pinnedEvaluate } from '../pinned-evaluate';

interface Comment { id: number; author: string; done?: boolean; parentId?: number }
interface WorkerCommentsProbe {
  editor: {
    whenLayoutComplete(options: { timeoutMs: number }): Promise<number>;
    getComments(): Comment[];
    scrollToComment(id: number): Promise<boolean>;
    scrollToChange(id: number): Promise<boolean>;
    insertCommentReply(id: number, text: string, author: string): Promise<number | null>;
    resolveComment(id: number): void;
  } | null;
  errors: string[];
  load(bytes: number[]): void;
  setComments(comments: unknown[] | undefined): void;
  pageInView(index: number): boolean;
  scrollToTop(): void;
  saved(): Promise<number[]>;
}
declare global { interface Window { __workerComments: WorkerCommentsProbe } }

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

async function reviewDocx(spanning = false, reply = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${W}><w:body>` +
    '<w:p w14:paraId="00000001"><w:r><w:t>Intro</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>' +
    '<w:p w14:paraId="00000002"><w:r><w:t>Middle</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>' +
    '<w:p w14:paraId="00000003"><w:commentRangeStart w:id="3"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p>' +
    '<w:p w14:paraId="00000004"><w:r><w:t>Before </w:t></w:r><w:ins w:id="9" w:author="A" w:date="2026-10-01T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>' +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  if (spanning) zip.file('word/document.xml', `<w:document ${W}><w:body>` + Array.from({ length: 7 }, (_, index) =>
    `<w:p w14:paraId="${String(index + 1).padStart(8, '0')}">` + (index === 3 ? '<w:commentRangeStart w:id="3"/>' : '') +
    `<w:r><w:t>Page ${index + 1}</w:t></w:r>` + (index === 6 ? '<w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>' : '<w:r><w:br w:type="page"/></w:r>') + '</w:p>'
  ).join('') + '<w:sectPr/></w:body></w:document>');
  if (reply) {
    zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/></Types>'));
    zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('</Relationships>', '<Relationship Id="rIdExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>'));
    zip.file('word/commentsExtended.xml', '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="0000000B" w15:paraIdParent="0000000A"/></w15:commentsEx>');
  }
  zip.file('word/comments.xml', `<w:comments ${W}><w:comment w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
    '<w:p w14:paraId="0000000A"><w:r><w:t>Parent note</w:t></w:r></w:p></w:comment>' +
    (reply ? '<w:comment w:id="4" w:author="B"><w:p w14:paraId="0000000B"><w:r><w:t>Imported reply</w:t></w:r></w:p></w:comment>' : '') + '</w:comments>');
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

async function open(page: Page, spanning = false, reply = false, hosted?: unknown[]) {
  await page.goto('/docx-worker-comments.html');
  if (hosted) await pinnedEvaluate(page, (comments) => window.__workerComments.setComments(comments), hosted);
  await pinnedEvaluate(page, (bytes) => window.__workerComments.load(bytes), await reviewDocx(spanning, reply));
  await expect.poll(() => page.evaluate(() => !!window.__workerComments.editor), { timeout: 60_000 }).toBe(true);
  expect(await pinnedEvaluate(page, () => window.__workerComments.editor!.whenLayoutComplete({ timeoutMs: 60_000 }))).toBe(spanning ? 7 : 3);
}

const thirdPageInView = (page: Page) => page.evaluate(() => window.__workerComments.pageInView(2));
const comments = (page: Page) => page.evaluate(() => window.__workerComments.editor!.getComments());

test('worker editor ref navigates to, replies to and resolves comments, and saves them', async ({ page }) => {
  test.setTimeout(180_000);
  await open(page);
  expect((await comments(page)).map(({ id, done }) => ({ id, done: done ?? false }))).toEqual([{ id: 3, done: false }]);
  expect(await thirdPageInView(page)).toBe(false);
  expect(await pinnedEvaluate(page, () => window.__workerComments.editor!.scrollToComment(3))).toBe(true);
  await expect.poll(() => thirdPageInView(page)).toBe(true);
  await page.evaluate(() => window.__workerComments.scrollToTop());
  await expect.poll(() => thirdPageInView(page)).toBe(false);
  expect(await pinnedEvaluate(page, () => window.__workerComments.editor!.scrollToChange(9))).toBe(true);
  await expect.poll(() => thirdPageInView(page)).toBe(true);
  expect(await pinnedEvaluate(page, () => window.__workerComments.editor!.scrollToComment(999))).toBe(false);
  expect(await pinnedEvaluate(page, () => window.__workerComments.editor!.scrollToChange(999))).toBe(false);
  expect(await pinnedEvaluate(page, () => window.__workerComments.editor!.insertCommentReply(999, 'Lost', 'Tester'))).toBeNull();
  const reply = await pinnedEvaluate(page, () => window.__workerComments.editor!.insertCommentReply(3, 'Reply text', 'Tester'));
  expect(typeof reply).toBe('number');
  expect((await comments(page)).find(comment => comment.id === reply)).toMatchObject({ parentId: 3, author: 'Tester' });
  await page.evaluate(() => window.__workerComments.editor!.resolveComment(3));
  await expect.poll(async () => (await comments(page)).find(comment => comment.id === 3)?.done).toBe(true);
  const zip = await JSZip.loadAsync(Uint8Array.from(await pinnedEvaluate(page, () => window.__workerComments.saved())));
  expect(await zip.file('word/comments.xml')!.async('string')).toContain('Reply text');
  const extended = await zip.file('word/commentsExtended.xml')?.async('string');
  expect(extended).toMatch(/w15:paraIdParent="[0-9A-F]{8}"/);
  expect(extended).toMatch(/w15:done="1"/);
  expect(await page.evaluate(() => window.__workerComments.errors)).toEqual([]);
});


test('worker review navigation reveals the start of a range across unbuilt pages', async ({ page }) => {
  test.setTimeout(180_000);
  await open(page, true);
  expect(await pinnedEvaluate(page, () => window.__workerComments.editor!.scrollToComment(3))).toBe(true);
  await expect.poll(() => page.evaluate(() => window.__workerComments.pageInView(3))).toBe(true);
  expect(await page.evaluate(() => window.__workerComments.pageInView(6))).toBe(false);
  expect(await page.evaluate(() => window.__workerComments.errors)).toEqual([]);
});

test('worker editor navigates and replies to an imported reply and immediately saves its resolution', async ({ page }) => {
  test.setTimeout(180_000);
  await open(page, false, true);
  expect((await comments(page)).find(comment => comment.id === 4)).toMatchObject({ parentId: 3 });
  expect(await pinnedEvaluate(page, () => window.__workerComments.editor!.scrollToComment(4))).toBe(true);
  await expect.poll(() => thirdPageInView(page)).toBe(true);
  const reply = await pinnedEvaluate(page, () => window.__workerComments.editor!.insertCommentReply(4, 'Nested reply', 'Tester'));
  expect(typeof reply).toBe('number');
  expect((await comments(page)).find(comment => comment.id === reply)).toMatchObject({ parentId: 4 });
  const bytes = await pinnedEvaluate(page, () => {
    window.__workerComments.editor!.resolveComment(4);
    return window.__workerComments.saved();
  });
  const zip = await JSZip.loadAsync(Uint8Array.from(bytes));
  expect(await zip.file('word/comments.xml')!.async('string')).toContain('Nested reply');
  expect(await zip.file('word/commentsExtended.xml')!.async('string')).toMatch(/w15:paraId="0000000B"[^>]*w15:done="1"/);
  expect(await page.evaluate(() => window.__workerComments.errors)).toEqual([]);
});

const note = (id: number, text: string) => ({ id, author: 'Host', date: '2026-02-01T00:00:00Z',
  content: [{ type: 'paragraph', formatting: {}, content: [{ type: 'run', formatting: {}, content: [{ type: 'text', text }] }] }] });

async function savedParts(page: Page) {
  const zip = await JSZip.loadAsync(Uint8Array.from(await pinnedEvaluate(page, () => window.__workerComments.saved())));
  return { comments: await zip.file('word/comments.xml')!.async('string'), document: await zip.file('word/document.xml')!.async('string') };
}

test('host comments given only through the controlled prop show in the sidebar, follow prop updates and survive save', async ({ page }) => {
  test.setTimeout(180_000);
  await open(page, false, false, [note(3, 'Host note'), note(8, 'Host only')]);
  const cards = page.locator('.docx-comment-card');
  await expect(cards.filter({ hasText: 'Host note' })).toBeVisible();
  await expect(cards.filter({ hasText: 'Host only' })).toBeVisible();
  await expect(cards.filter({ hasText: 'Parent note' })).toHaveCount(0);
  let saved = await savedParts(page);
  expect(saved.comments).toContain('Host note');
  expect(saved.comments).toContain('Host only');
  expect(saved.comments).not.toContain('Parent note');
  const anchored = saved.comments.match(/<w:comment [^>]*w:id="(\d+)"(?:(?!<\/w:comment>)[\s\S])*Host note/)?.[1];
  expect(anchored).toBeDefined();
  expect(saved.document).toContain(`<w:commentRangeStart w:id="${anchored}"`);
  await pinnedEvaluate(page, (comments) => window.__workerComments.setComments(comments), [note(3, 'Edited note')]);
  await expect(cards.filter({ hasText: 'Edited note' })).toBeVisible();
  await expect(cards.filter({ hasText: 'Host only' })).toHaveCount(0);
  saved = await savedParts(page);
  expect(saved.comments).toContain('Edited note');
  expect(saved.comments).not.toContain('Host only');
  expect(await page.evaluate(() => window.__workerComments.errors)).toEqual([]);
});
