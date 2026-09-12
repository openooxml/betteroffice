import { expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import type { DiagramHandle } from '@betteroffice/vsdx';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { Ribbon } from './Ribbon';
import { RibbonCommandsProvider } from './commands';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { fireEvent, render } = await import('@testing-library/react');

test('renders all tabs and supports click and roving arrow-key selection', () => {
  const diagram = { snapshot: () => ({ pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: [] }] }), canUndo: () => false, canRedo: () => false } as unknown as DiagramHandle;
  const view = render(<RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={null} onMutation={() => {}} onError={() => {}} onDownload={() => {}}><Ribbon t={createT(en)} /></RibbonCommandsProvider>);
  const home = view.getByRole('tab', { name: 'Home' }); const insert = view.getByRole('tab', { name: 'Insert' });
  expect(view.getAllByRole('tab')).toHaveLength(7); expect(home.getAttribute('aria-selected')).toBe('true'); expect(home.tabIndex).toBe(0); expect(insert.tabIndex).toBe(-1);
  fireEvent.click(insert); expect(insert.getAttribute('aria-selected')).toBe('true'); expect(insert.tabIndex).toBe(0);
  fireEvent.keyDown(insert, { key: 'ArrowRight' }); const design = view.getByRole('tab', { name: 'Design' }); expect(design.getAttribute('aria-selected')).toBe('true'); expect(document.activeElement).toBe(design);
});
