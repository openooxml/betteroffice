import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { createRef } from 'react';
import { buildMirrorPage, type DisplayPage } from '@betteroffice/docx/layout/render';
import { ContentControlWidgets } from './ContentControlWidgets';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('clicking an occluded header checkbox mirror changes nothing', () => {
  const page: DisplayPage = {
    pageIndex: 0, width: 200, height: 200,
    primitives: [{ kind: 'rect', x: 10, y: 10, w: 40, h: 20, fill: '#fff' }],
    header: {
      rId: 'rIdHeader', kind: 'header', y: 0, height: 40,
      primitives: [{
        kind: 'rect', x: 10, y: 10, w: 20, h: 20, fill: '#fff',
        inlineSdtWidget: { kind: 'checkbox', groupId: 'header', pos: 1, checked: false },
      }],
    },
  };
  const containerRef = createRef<HTMLDivElement>();
  const applyYrsValue = mock(() => true);
  render(
    <div ref={containerRef}>
      <ContentControlWidgets containerRef={containerRef} applyYrsValue={applyYrsValue} />
    </div>
  );
  const mirror = buildMirrorPage(page);
  containerRef.current!.appendChild(mirror);
  const checkbox = mirror.querySelector<HTMLElement>('.layout-inline-sdt-widget')!;
  fireEvent.click(checkbox);
  fireEvent.keyDown(checkbox, { key: ' ' });
  expect(applyYrsValue).not.toHaveBeenCalled();
  expect(checkbox.getAttribute('aria-checked')).toBe('false');

  page.primitives = [];
  const restored = buildMirrorPage(page);
  mirror.replaceWith(restored);
  fireEvent.click(restored.querySelector<HTMLElement>('.layout-inline-sdt-widget')!);
  expect(applyYrsValue).toHaveBeenCalledTimes(1);
  expect(applyYrsValue).toHaveBeenCalledWith(1, { kind: 'checkbox', checked: true }, undefined);
});
