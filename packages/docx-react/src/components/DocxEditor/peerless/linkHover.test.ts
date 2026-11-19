import { expect, test } from 'bun:test';
import { popupWithLink } from './linkHover';

const hit = { href: 'https://example.com/', displayText: '比亚迪', position: { top: 1, left: 2 } };

test('the popup shows the whole link text, not the run that was clicked', () => {
  expect(popupWithLink(hit, { url: 'https://example.com/', displayText: '比亚迪 2025 年度简报' }))
    .toEqual({ ...hit, displayText: '比亚迪 2025 年度简报' });
});

test('the clicked run stands when the link cannot be read', () => {
  expect(popupWithLink(hit, null)).toBe(hit);
  expect(popupWithLink(hit, { displayText: 'not a link' })).toBe(hit);
  expect(popupWithLink({ ...hit, tooltip: 'From paint' }, { url: 'https://example.com/', displayText: '' }))
    .toEqual({ ...hit, tooltip: 'From paint' });
});
