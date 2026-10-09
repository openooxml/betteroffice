import { expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createT, en } from '@betteroffice/vsdx-i18n';
import type { PageLayer } from '@betteroffice/vsdx';
import { LayersPanel } from './LayersPanel';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const t = createT(en);

const layers: PageLayer[] = [
  { index: 0, name: 'Trussing', visible: true, print: true, lock: false, active: false, color: '255', status: '0' },
  { index: 1, name: 'Lighting', visible: false, print: true, lock: false, active: false, color: '255', status: '0' },
];

test('lists layers with a visibility checkbox each', () => {
  const toggles: Array<[number, boolean]> = [];
  const view = render(
    <LayersPanel layers={layers} collapsed={false} onToggleCollapsed={() => {}} onToggleLayer={(index, visible) => toggles.push([index, visible])} t={t} />,
  );
  try {
    const boxes = view.container.querySelectorAll('input[type="checkbox"]');
    expect(boxes).toHaveLength(2);
    expect((boxes[0] as HTMLInputElement).checked).toBe(true);
    expect((boxes[1] as HTMLInputElement).checked).toBe(false);
    expect(view.getByText('Trussing')).toBeDefined();
    expect(view.getByText('Lighting')).toBeDefined();
    act(() => { fireEvent.click(boxes[1] as HTMLInputElement); });
    expect(toggles).toEqual([[1, true]]);
  } finally { view.unmount(); cleanup(); }
});

test('falls back to a numbered name and reports an empty page', () => {
  const unnamed: PageLayer[] = [{ ...layers[0], name: '' }];
  const view = render(
    <LayersPanel layers={unnamed} collapsed={false} onToggleCollapsed={() => {}} onToggleLayer={() => {}} t={t} />,
  );
  try {
    expect(view.getByText('Layer 0')).toBeDefined();
  } finally { view.unmount(); cleanup(); }
  const empty = render(
    <LayersPanel layers={[]} collapsed={false} onToggleCollapsed={() => {}} onToggleLayer={() => {}} t={t} />,
  );
  try {
    expect(empty.getByText('No layers on this page')).toBeDefined();
  } finally { empty.unmount(); cleanup(); }
});

test('collapses the layer list', () => {
  let collapsed = false;
  const view = render(
    <LayersPanel layers={layers} collapsed={collapsed} onToggleCollapsed={() => { collapsed = true; }} onToggleLayer={() => {}} t={t} />,
  );
  try {
    expect(view.container.querySelectorAll('input[type="checkbox"]')).toHaveLength(2);
    act(() => { fireEvent.click(view.getByTitle('Collapse Layers')); });
    expect(collapsed).toBe(true);
  } finally { view.unmount(); cleanup(); }
  const hidden = render(
    <LayersPanel layers={layers} collapsed onToggleCollapsed={() => {}} onToggleLayer={() => {}} t={t} />,
  );
  try {
    expect(hidden.container.querySelectorAll('input[type="checkbox"]')).toHaveLength(0);
  } finally { hidden.unmount(); cleanup(); }
});
