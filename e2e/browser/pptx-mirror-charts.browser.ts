import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { test, expect, type Page } from 'playwright/test';

const deck = fileURLToPath(new URL('../../crates/pptx-render/tests/fixtures/data-labels.pptx', import.meta.url));
const summary = /2 series, 3 categories$/;

async function open(page: Page) {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'charts.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: await readFile(deck),
  });
  await expect(page.getByTestId('pptx-worker-editor')).toBeVisible();
  await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
  await expect(page.getByTestId('pptx-slide-canvas')).toHaveAttribute('data-pptx-input-ready', 'true', { timeout: 60_000 });
}

type AXNode = { nodeId: string; ignored?: boolean; role?: { value?: unknown }; name?: { value?: unknown }; childIds?: string[] };
const tableRoles = new Set(['table', 'row', 'columnheader', 'rowheader', 'cell']);

/** The figure's accessibility subtree in tree order, as role and name; rows by role only. */
async function figureTree(page: Page) {
  const cdp = await page.context().newCDPSession(page);
  const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
  await cdp.detach();
  const byId = new Map(nodes.map((node) => [node.nodeId, node]));
  const out: string[][] = [];
  const visit = (node: AXNode | undefined) => {
    if (!node) return;
    const role = String(node.role?.value);
    if (!node.ignored && (role === 'figure' || tableRoles.has(role))) {
      out.push(role === 'row' ? ['row'] : [role, String(node.name?.value ?? '')]);
    }
    for (const id of node.childIds ?? []) visit(byId.get(id));
  };
  visit(nodes.find((node) => !node.ignored && node.role?.value === 'figure'));
  return out;
}

test('pptx mirror charts expose their values as a table inside a named figure', async ({ page }) => {
  await open(page);
  const mirror = page.getByTestId('pptx-slide-mirror');
  const figure = mirror.getByRole('figure', { name: summary });
  await expect(figure).toBeAttached();
  const table = figure.getByRole('table');
  await expect(table).toHaveAttribute('aria-rowcount', '4');
  await expect(table).toHaveAttribute('aria-colcount', '3');
  await expect(table.getByRole('columnheader')).toHaveText(['North', 'South']);
  await expect(table.getByRole('rowheader')).toHaveText(['Q1', 'Q2', 'Q3']);
  await expect(table.getByRole('cell')).toHaveText(['', '12', '8', '19', '14', '7', '21']);
  await expect(mirror.getByRole('img', { name: summary })).toHaveCount(0);
  await expect.poll(async () => (await figureTree(page)).map(([role, name]) => role === 'figure' ? [role, summary.test(name!)] : [role, name]))
    .toEqual([
      ['figure', true],
      ['table', ''],
      ['row'], ['cell', ''], ['columnheader', 'North'], ['columnheader', 'South'],
      ['row'], ['rowheader', 'Q1'], ['cell', '12'], ['cell', '8'],
      ['row'], ['rowheader', 'Q2'], ['cell', '19'], ['cell', '14'],
      ['row'], ['rowheader', 'Q3'], ['cell', '7'], ['cell', '21'],
    ]);
});
