import { expect, test, type Page } from 'playwright/test';
import type { DocxEngineWorkerResponse } from '../../packages/docx/src/yrs/residentEngineWorkerProtocol';
import { pagedDocx } from '../../packages/docx-react/src/components/DocxEditor/__fixtures__/pagedDocx';
import { pinnedEvaluate } from './pinned-evaluate';

interface RecoveryProbe {
  worker?: Worker;
  incarnation: number;
  version: number;
  checkpoints: number[];
}

declare global {
  interface Window { __pasteRecovery: RecoveryProbe }
}

async function observeWorker(page: Page) {
  await page.addInitScript(() => {
    const probe: RecoveryProbe = { incarnation: 0, version: 0, checkpoints: [] };
    window.__pasteRecovery = probe;
    window.Worker = class extends Worker {
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        this.addEventListener('message', ({ data }: MessageEvent<DocxEngineWorkerResponse>) => {
          if (!data.ok) return;
          if (data.commandInitial) {
            probe.worker = this;
            probe.incarnation = data.commandInitial.incarnation;
            probe.version = data.commandInitial.version;
          }
          if (this !== probe.worker) return;
          if (data.commandAck) probe.version = Math.max(probe.version, data.commandAck.version);
          if (data.checkpoint) probe.checkpoints.push(data.checkpoint.version);
        });
      }
    };
  });
}

for (const scenario of ['PASTE_ON_CHECKPOINT', 'PASTE_BEFORE_CHECKPOINT', 'CONTROL_NO_PASTE'] as const) {
  test(`DOCX worker crash preserves edits in ${scenario}`, async ({ page, context }) => {
    await observeWorker(page);
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await page.goto('/?format=docx&peerless=1');
    await page.locator('input[type=file]').last().setInputFiles({ name: 'paste-recovery.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(await pagedDocx(1)) });
    const input = page.locator('.paged-editor__input');
    const mirror = page.locator('.canvas-page-mirror').first();
    await expect(mirror).toContainText('Page 1');
    await input.press('ControlOrMeta+End');
    const version = () => page.evaluate(() => window.__pasteRecovery.version);
    const reach = (next: number) => page.waitForFunction(next => window.__pasteRecovery.version >= next, next);
    const boundary = scenario === 'PASTE_BEFORE_CHECKPOINT' ? 126 : 127;
    let prefix = 'Page 1';
    for (let current = await version(); current < boundary; current = await version()) {
      await page.keyboard.insertText('i');
      prefix += 'i';
      await reach(current + 1);
    }
    expect(await version()).toBe(boundary);
    if (scenario === 'CONTROL_NO_PASTE') {
      await page.keyboard.insertText('x');
      prefix += 'x';
      await reach(128);
    } else {
      await pinnedEvaluate(page, () => navigator.clipboard.write([new ClipboardItem({
        'text/plain': new Blob(['Rich'], { type: 'text/plain' }),
        'text/html': new Blob(['<p><b>Rich</b></p>'], { type: 'text/html' }),
      })]));
      await input.press('ControlOrMeta+v');
      prefix += 'Rich';
      const control = page.getByRole('button', { name: 'Paste Options (Ctrl)' });
      const choose = async (name: string) => {
        const before = await version();
        await control.click();
        await page.getByRole('menuitemradio', { name }).click();
        await reach(before + 1);
      };
      await expect(control).toBeVisible();
      expect(await version()).toBe(boundary + 1);
      if (scenario === 'PASTE_BEFORE_CHECKPOINT') await choose('Merge Formatting');
      expect(await version()).toBe(128);
      await choose('Keep Text Only');
      expect(await page.evaluate(() => window.__pasteRecovery.checkpoints)).toEqual([]);
    }
    await page.keyboard.insertText('AfterOpt');
    const expected = prefix + 'AfterOpt';
    await expect(mirror).toContainText(expected);
    await page.waitForFunction(() => window.__pasteRecovery.checkpoints.length > 0);
    await page.evaluate(() => {
      const worker = window.__pasteRecovery.worker!;
      worker.dispatchEvent(new ErrorEvent('error', { message: 'Browser worker failure', error: new Error('Browser worker failure') }));
      worker.terminate();
    });
    await page.waitForFunction(() => window.__pasteRecovery.incarnation === 2);
    await expect(mirror).toContainText(expected);
    await input.focus();
    await page.keyboard.insertText('Recovered');
    await expect(mirror).toContainText(expected + 'Recovered');
    await expect(page.getByText(/The file may be damaged|Couldn't open this document/)).toHaveCount(0);
    if (scenario !== 'CONTROL_NO_PASTE') {
      await input.press('ControlOrMeta+a');
      await input.press('ControlOrMeta+c');
      await expect(async () => {
        const html = await pinnedEvaluate(page, async () => {
          const item = (await navigator.clipboard.read()).find(item => item.types.includes('text/html'));
          return item ? (await item.getType('text/html')).text() : '';
        });
        expect(html).toContain('Recovered');
        expect(html).not.toMatch(/font-weight:\s*(bold|[7-9]00)/i);
      }).toPass({ timeout: 15_000 });
    }
  });
}
