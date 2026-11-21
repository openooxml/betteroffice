import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { useRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { render } from '@testing-library/react';
import { useToolbarOverflow } from './useToolbarOverflow';

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function rect(width: number): DOMRect {
  return {
    width,
    height: 28,
    top: 0,
    left: 0,
    right: width,
    bottom: 28,
    x: 0,
    y: 0,
    toJSON() {},
  } as DOMRect;
}

function Toolbar() {
  const items = useRef<HTMLDivElement>(null);
  const more = useRef<HTMLSpanElement>(null);
  const overflow = useToolbarOverflow({
    items,
    more,
    canHide: (unit) => unit.dataset.fixed !== 'true',
  });

  return (
    <div ref={(element) => {
      if (element) element.getBoundingClientRect = () => rect(120);
    }}>
      <div ref={items}>
        <span data-testid="first" ref={(element) => {
          if (element) element.getBoundingClientRect = () => rect(56);
        }} />
        <span data-testid="second" ref={(element) => {
          if (element) element.getBoundingClientRect = () => rect(50);
        }} />
        <span data-fixed="true" ref={(element) => {
          if (element) element.getBoundingClientRect = () => rect(35);
        }} />
      </div>
      {overflow.hidden.length > 0 && (
        <span ref={(element) => {
          more.current = element;
          if (element) element.getBoundingClientRect = () => rect(28);
        }} />
      )}
    </div>
  );
}

test('remeasures after the More trigger mounts', () => {
  const view = render(<Toolbar />);
  expect(view.getByTestId('first').getAttribute('aria-hidden')).toBeNull();
  expect(view.getByTestId('second').getAttribute('aria-hidden')).toBe('true');
});
