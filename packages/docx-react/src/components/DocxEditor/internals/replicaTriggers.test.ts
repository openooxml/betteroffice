import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { pagePressNeedsReplica } from './replicaTriggers';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function press(target: Element, pointerType: string): boolean {
  let needs = false;
  target.addEventListener('pointerdown', (event) => {
    needs = pagePressNeedsReplica(event as PointerEvent);
  }, { once: true });
  target.dispatchEvent(Object.assign(new Event('pointerdown', { bubbles: true }), { pointerType }));
  return needs;
}

test.each([
  ['mouse', true],
  ['', true],
  ['touch', false],
  ['pen', false],
])('a %s press on a page needs the replica: %s', (pointerType, expected) => {
  const page = document.createElement('canvas');
  page.className = 'canvas-page';
  document.body.append(page);
  try {
    expect(press(page, pointerType)).toBe(expected);
  } finally {
    page.remove();
  }
});

test('a press on an overlay outside the pages does not need the replica', () => {
  const overlay = document.createElement('button');
  document.body.append(overlay);
  try {
    expect(press(overlay, 'mouse')).toBe(false);
  } finally {
    overlay.remove();
  }
});
