import { expect, test } from 'bun:test';
import { createPptxCommandController } from '../commands/createPptxCommandStore';
import { testBinding } from '../commands/testing';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import { pptxContextMenuEntries, pptxContextMenuSections, type PptxContextMenuTarget } from './menus';

function menu(target: PptxContextMenuTarget, overrides: Parameters<typeof testBinding>[0] = {}) {
  const harness = testBinding(overrides);
  const controller = createPptxCommandController();
  controller.attach(harness.binding);
  return { harness, entries: pptxContextMenuEntries(target, controller.store, harness.state.env.translate),
    sections: pptxContextMenuSections(target, controller.store, harness.state.env.translate) };
}

function labels(entries: readonly OverflowMenuEntry[]): string[] {
  return entries.map((entry) => entry.kind === 'separator' ? '|' : entry.kind === 'group' ? entry.id : entry.label);
}

test('text and table menus follow Office clipboard and table order', () => {
  expect(labels(menu('text').entries)).toEqual(['Cut', 'Copy', 'Paste']);
  const table = menu('table');
  expect(table.sections.map((section) => section.id)).toEqual(['clipboard', 'table']);
  expect(labels(table.entries)).toEqual(['Cut', 'Copy', 'Paste', '|', 'Insert', 'Delete']);
  const insert = table.entries[4];
  const remove = table.entries[5];
  expect(insert.kind).toBe('submenu');
  expect(remove.kind).toBe('submenu');
  if (insert.kind === 'submenu') expect(labels(insert.entries)).toEqual([
    'Insert rows above', 'Insert rows below', 'Insert columns to the left', 'Insert columns to the right',
  ]);
  if (remove.kind === 'submenu') expect(labels(remove.entries)).toEqual(['Delete rows', 'Delete columns', 'Delete table']);
});

test('shape and slide menus have the expected sections and order', () => {
  const shape = menu('shape');
  expect(shape.sections.map((section) => section.id)).toEqual(['object', 'arrange']);
  expect(labels(shape.entries)).toEqual(['Delete', '|', 'Bring to front', 'Send to back']);
  const front = shape.entries[2];
  const back = shape.entries[3];
  if (front.kind === 'submenu') expect(labels(front.entries)).toEqual(['Bring to front', 'Bring forward']);
  if (back.kind === 'submenu') expect(labels(back.entries)).toEqual(['Send to back', 'Send backward']);
  expect(labels(menu('slide').entries)).toEqual(['New slide', 'Duplicate slide', 'Delete slide']);
  expect(menu('empty').entries).toEqual([]);
});

test('items show command shortcuts and route every menu action through the store', async () => {
  for (const target of ['text', 'shape', 'table', 'slide'] as const) {
    const { entries, harness } = menu(target);
    const leaves = (items: readonly OverflowMenuEntry[]): Extract<OverflowMenuEntry, { kind: 'item' }>[] =>
      items.flatMap((entry) => entry.kind === 'item' ? [entry]
        : entry.kind === 'submenu' || entry.kind === 'group' ? leaves(entry.entries) : []);
    for (const entry of leaves(entries)) entry.onSelect();
    await Promise.resolve();
    expect(harness.calls.map(({ id, args }) => `${id}:${JSON.stringify(args)}`)).toEqual(leaves(entries).map((entry) => entry.id));
    if (target === 'text') expect(leaves(entries).map((entry) => entry.shortcut)).toEqual([
      expect.stringMatching(/Ctrl\+X|⌘X/), expect.stringMatching(/Ctrl\+C|⌘C/), expect.stringMatching(/Ctrl\+V|⌘V/),
    ]);
  }
});

test('disabled commands and submenus preserve their reason', () => {
  const text = menu('text', { clipboardRead: false, clipboardWrite: false });
  for (const entry of text.entries) {
    expect(entry.kind === 'item' && entry.disabled).toBe(true);
    expect(entry.kind === 'item' && entry.description).toBe('Clipboard access is unavailable or was denied.');
  }
  const table = menu('table', { table: null });
  expect(table.entries[4]).toMatchObject({ kind: 'submenu', disabled: true });
  expect(table.entries[5]).toMatchObject({ kind: 'submenu', disabled: true });
  expect(menu('slide', { slideCount: 1 }).entries[2]).toMatchObject({ disabled: true,
    description: 'The last slide cannot be deleted.' });
  const top = menu('shape', { shape: { ...testBinding().state.env.shape!, order: { index: 2, count: 3 } } });
  expect(top.entries[2]).toMatchObject({ kind: 'submenu', disabled: true });
  expect(top.entries[3]).toMatchObject({ kind: 'submenu', disabled: false });
  const picture = menu('shape', { shape: { formattable: false, geometry: null, fill: null, stroke: null,
    strokeWidth: null, adjustments: {}, order: { index: 0, count: 2 } } });
  expect(labels(picture.entries)).toEqual(['Delete', '|', 'Bring to front', 'Send to back']);
  expect(picture.entries[0]).toMatchObject({ disabled: false });
  expect(picture.entries[2]).toMatchObject({ disabled: false });
  expect(picture.entries[3]).toMatchObject({ disabled: true });
  expect(menu('shape', { shape: null }).entries[0]).toMatchObject({ disabled: true });
  for (const target of ['shape', 'slide'] as const) {
    expect(menu(target, { readOnly: true }).entries.filter((entry) => entry.kind !== 'separator'))
      .toEqual(expect.arrayContaining([expect.objectContaining({ disabled: true })]));
  }
  const readOnly = menu('table', { readOnly: true }).entries;
  expect(readOnly[0]).toMatchObject({ disabled: true });
  expect(readOnly[1]).toMatchObject({ disabled: false });
  expect(readOnly[2]).toMatchObject({ disabled: true });
  expect(readOnly[4]).toMatchObject({ disabled: true });
  expect(readOnly[5]).toMatchObject({ disabled: true });
});
