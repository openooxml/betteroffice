import { test, expect } from 'playwright/test';

const INVALID_DOCX = {
  name: 'not-a-document.docx',
  mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  buffer: Buffer.from('This is not a DOCX archive'),
};

test('docx demo: initial fetch errors stay on the load-error path', async ({ page }) => {
  await page.route('**/betteroffice-demo.docx', (route) => route.abort());
  await page.goto('/docx');

  await expect(
    page.getByRole('alert').filter({ hasText: 'Failed to load the demo document' })
  ).toContainText('Failed to load the demo document');
  await expect(page.getByTestId('docx-editor')).toHaveCount(0);
});

test('docx demo: runtime errors are dismissible without unmounting the editor', async ({ page }) => {
  await page.goto('/docx');
  const editor = page.getByTestId('docx-editor');
  await expect(editor).toBeVisible({ timeout: 60_000 });

  await page
    .locator('input[type="file"][accept^=".docx"]')
    .setInputFiles(INVALID_DOCX);

  const errorBanner = page.getByRole('alert').filter({ hasText: 'Editor error:' });
  await expect(errorBanner).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('docx-editor')).toBeVisible();

  await errorBanner.getByRole('button', { name: 'Dismiss editor error' }).click();
  await expect(errorBanner).toBeHidden();
  await expect(page.getByTestId('docx-editor')).toBeVisible();
});
