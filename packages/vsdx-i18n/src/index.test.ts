import { expect, test } from 'bun:test';
import { createT, diagnosticMessage, en } from './index';

const t = createT(en);

test('maps diagnostic codes to localized messages', () => {
  expect(diagnosticMessage(t, 'integrity', 'missing-media')).toBe('An image could not be recovered.');
});

test('degrades unknown diagnostic codes gracefully', () => {
  expect(diagnosticMessage(t, 'fidelity', 'new-code')).toBe('fidelity notice: new-code');
});
