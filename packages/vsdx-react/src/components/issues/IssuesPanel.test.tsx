import { expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createT, en } from '@betteroffice/vsdx-i18n';
import type { ValidationIssue } from '@betteroffice/vsdx';
import { IssuesPanel } from './IssuesPanel';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { cleanup, fireEvent, render } = await import('@testing-library/react');
const t = createT(en);

const issues: ValidationIssue[] = [
  { id: 'dangling-connector:visio/pages/page1.xml:4:', rule: 'dangling-connector', severity: 'warning', pageId: 'page:1', shapeId: 'shape:4', otherShapeId: null, endpoint: 'end', row: null },
  { id: 'overlapping-shapes:visio/pages/page1.xml:1:2', rule: 'overlapping-shapes', severity: 'warning', pageId: 'page:1', shapeId: 'shape:1', otherShapeId: 'shape:2', endpoint: null, row: null },
  { id: 'empty-shape-data:visio/pages/page1.xml:1:Owner', rule: 'empty-shape-data', severity: 'error', pageId: 'page:1', shapeId: 'shape:1', otherShapeId: null, endpoint: null, row: 'Owner' },
];

const names: Record<string, string> = { 'shape:1': 'Box', 'shape:2': 'Other' };

test('names every rule and reports the shape a row selects', () => {
  const selected: string[] = [];
  const view = render(
    <IssuesPanel
      issues={issues}
      collapsed={false}
      onToggleCollapsed={() => {}}
      onSelect={(issue) => selected.push(issue.id)}
      getShapeName={(_pageId, shapeId) => names[shapeId] ?? null}
      t={t}
    />,
  );
  const rows = [...view.container.querySelectorAll('li button')] as HTMLButtonElement[];
  expect(rows.map((row) => row.textContent)).toEqual([
    'Connector Shape shape:4 is glued at only one end (end)',
    'Shape Box overlaps Other',
    'Shape Box has an empty required field (Owner)',
  ]);
  expect(view.container.querySelector('header span span')?.textContent).toBe('3');
  fireEvent.click(rows[1]);
  expect(selected).toEqual(['overlapping-shapes:visio/pages/page1.xml:1:2']);
  cleanup();
});

test('a collapsed panel hides the list and an empty one says so', () => {
  const collapsed = render(
    <IssuesPanel issues={issues} collapsed onToggleCollapsed={() => {}} onSelect={() => {}} getShapeName={() => null} t={t} />,
  );
  expect(collapsed.container.querySelector('li')).toBeNull();
  cleanup();
  const empty = render(
    <IssuesPanel issues={[]} collapsed={false} onToggleCollapsed={() => {}} onSelect={() => {}} getShapeName={() => null} t={t} />,
  );
  expect(empty.container.querySelector('p')?.textContent).toBe('No issues on this page');
  cleanup();
});
