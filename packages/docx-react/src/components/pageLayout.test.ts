import { describe, expect, test } from 'bun:test';
import { createT, en } from '@betteroffice/docx-i18n';
import { formatLayoutLength, formatLayoutMargins, MARGIN_PRESETS, PAGE_SIZES, pageOrientationPatch } from './pageLayout';

const t = createT(en, 'en');

describe('page layout presets', () => {
  test('uses Word margin values in twips', () => {
    expect(MARGIN_PRESETS.map(({ labelKey, ...margins }) => [labelKey, margins])).toEqual([
      ['layout.normal', { marginTop: 1440, marginBottom: 1440, marginLeft: 1440, marginRight: 1440 }],
      ['layout.narrow', { marginTop: 720, marginBottom: 720, marginLeft: 720, marginRight: 720 }],
      ['layout.moderate', { marginTop: 1440, marginBottom: 1440, marginLeft: 1080, marginRight: 1080 }],
      ['layout.wide', { marginTop: 1440, marginBottom: 1440, marginLeft: 2880, marginRight: 2880 }],
    ]);
  });

  test('shares the dialog paper sizes in portrait twips', () => {
    expect(PAGE_SIZES.map(({ width, height }) => [width, height])).toEqual([
      [12240, 15840], [11906, 16838], [12240, 20160], [16838, 23811],
      [8391, 11906], [9979, 14175], [10440, 15120],
    ]);
  });

  test('formats secondary margin text in inches and centimeters', () => {
    expect(formatLayoutMargins(MARGIN_PRESETS[0], 'inch', t)).toBe('Top 1 in  Bottom 1 in  Left 1 in  Right 1 in');
    expect(formatLayoutMargins(MARGIN_PRESETS[0], 'cm', t)).toBe('Top 2.54 cm  Bottom 2.54 cm  Left 2.54 cm  Right 2.54 cm');
    expect(formatLayoutMargins(MARGIN_PRESETS[2], 'inch', t)).toBe('Top 1 in  Bottom 1 in  Left 0.75 in  Right 0.75 in');
    expect(formatLayoutMargins(MARGIN_PRESETS[1], 'cm', t)).toBe('Top 1.27 cm  Bottom 1.27 cm  Left 1.27 cm  Right 1.27 cm');
  });

  test('formats paper dimensions in either ruler unit and the active locale', () => {
    expect(formatLayoutLength(12240, 'inch', t)).toBe('8.5 in');
    expect(formatLayoutLength(12240, 'cm', t)).toBe('21.59 cm');
    expect(formatLayoutLength(1440, 'cm', t, 'de')).toBe('2,54 cm');
  });

  test('swaps custom dimensions only when orientation changes', () => {
    const properties = { pageWidth: 10000, pageHeight: 20000 };
    expect(pageOrientationPatch(properties, 'landscape')).toEqual({ orientation: 'landscape', pageWidth: 20000, pageHeight: 10000 });
    expect(pageOrientationPatch(properties, 'portrait')).toEqual({ orientation: 'portrait' });
    expect(pageOrientationPatch({ pageWidth: 20000, pageHeight: 10000 }, 'portrait')).toEqual({ orientation: 'portrait', pageWidth: 10000, pageHeight: 20000 });
  });
});
