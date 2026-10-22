import { expect, test } from 'bun:test';
import { isGoogleFontsEnabled } from '@betteroffice/docx/utils/fontLoader';

test('unit tests run without the automatic Google Fonts lookup', () => {
  expect(isGoogleFontsEnabled()).toBe(false);
});
