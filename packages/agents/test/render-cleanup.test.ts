import { expect, mock, test } from 'bun:test';
import { openDocx, type DocumentRenderer } from '../src';
import { cleanupGuard } from '../src/cleanup';
import { renderDocxPage } from '../src/render';
import { fixture } from './fixture';

const breakingDestroy = (session: Parameters<DocumentRenderer>[0], error: Error) => {
  const destroy = session.destroy.bind(session);
  session.destroy = () => { destroy(); throw error; };
};

test('a render failure survives a cleanup that also throws', async () => {
  const original = new Error('render failed');
  const secondary = new Error('attempted to take ownership of Rust value while it was borrowed');
  const doc = await openDocx(await fixture(), {
    renderer: async (session) => { breakingDestroy(session, secondary); throw original; },
  });
  try {
    const failure = await doc.render(1).catch((error: unknown) => error);
    expect(failure).toBe(original);
    expect((failure as { suppressed?: unknown }).suppressed).toBe(secondary);
  } finally { doc.close(); }
});

test('a cleanup failure after a successful render is still reported', async () => {
  const secondary = new Error('free failed');
  const doc = await openDocx(await fixture(), {
    renderer: async (session, page) => { breakingDestroy(session, secondary); return renderDocxPage(session, page); },
  });
  try { await expect(doc.render(1)).rejects.toBe(secondary); }
  finally { doc.close(); }
});

test('a cleanup guard cleans up once and leaves a frozen failure alone', () => {
  const cleanup = mock(() => { throw new Error('free failed'); });
  const guard = cleanupGuard(cleanup);
  const frozen = Object.freeze(new Error('frozen'));
  expect(guard.failed(frozen)).toBe(frozen);
  guard.run();
  expect(cleanup).toHaveBeenCalledTimes(1);
  expect(() => cleanupGuard(cleanup).run()).toThrow('free failed');
});

test('a cleanup guard keeps a failure whose suppressed property cannot be written', () => {
  const fixed = Object.defineProperty(new Error('fixed'), 'suppressed', { value: 'prior' });
  const trapped = new Proxy(new Error('trapped'), { defineProperty() { throw new Error('trap'); } });
  const revoked = Proxy.revocable(new Error('revoked'), {});
  revoked.revoke();
  for (const failure of [fixed, trapped, revoked.proxy]) {
    expect(cleanupGuard(() => { throw new Error('free failed'); }).failed(failure)).toBe(failure);
  }
  expect((fixed as Error & { suppressed: unknown }).suppressed).toBe('prior');
});
