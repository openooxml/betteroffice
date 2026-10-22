import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { createElement, useEffect } = await import('react');
const { createRoot } = await import('react-dom/client');

test.if(ownsDom)('React work still scheduled at DOM teardown runs before the globals are removed', async () => {
  const container = document.body.appendChild(document.createElement('div'));
  let effects = 0;
  const Probe = () => {
    useEffect(() => { effects++; }, []);
    return null;
  };
  createRoot(container).render(createElement(Probe));
  expect(effects).toBe(0);
  await GlobalRegistrator.unregister();
  expect(effects).toBe(1);
});
