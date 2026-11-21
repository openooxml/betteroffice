import type { TFunction } from '@betteroffice/docx-i18n';
import type { SectionProperties } from '@betteroffice/docx/types/document';

export const PAGE_SIZES = [
  { labelKey: 'dialogs.pageSetup.pageSizes.letter' as const, width: 12240, height: 15840 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a4' as const, width: 11906, height: 16838 },
  { labelKey: 'dialogs.pageSetup.pageSizes.legal' as const, width: 12240, height: 20160 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a3' as const, width: 16838, height: 23811 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a5' as const, width: 8391, height: 11906 },
  { labelKey: 'dialogs.pageSetup.pageSizes.b5' as const, width: 9979, height: 14175 },
  { labelKey: 'dialogs.pageSetup.pageSizes.executive' as const, width: 10440, height: 15120 },
] as const;

export const MARGIN_PRESETS = [
  { labelKey: 'layout.normal', marginTop: 1440, marginBottom: 1440, marginLeft: 1440, marginRight: 1440 },
  { labelKey: 'layout.narrow', marginTop: 720, marginBottom: 720, marginLeft: 720, marginRight: 720 },
  { labelKey: 'layout.moderate', marginTop: 1440, marginBottom: 1440, marginLeft: 1080, marginRight: 1080 },
  { labelKey: 'layout.wide', marginTop: 1440, marginBottom: 1440, marginLeft: 2880, marginRight: 2880 },
] as const;

export type PageMargins = Pick<SectionProperties, 'marginTop' | 'marginBottom' | 'marginLeft' | 'marginRight'>;

export function formatLayoutLength(twips: number, unit: 'inch' | 'cm', t: TFunction, lang = 'en'): string {
  const value = twips / 1440 * (unit === 'cm' ? 2.54 : 1);
  return `${new Intl.NumberFormat(lang, { maximumFractionDigits: 2 }).format(value)} ${t(unit === 'cm' ? 'common.units.cm' : 'common.units.inch')}`;
}

export function formatLayoutMargins(margins: PageMargins, unit: 'inch' | 'cm', t: TFunction, lang = 'en'): string {
  return [
    `${t('dialogs.pageSetup.top')} ${formatLayoutLength(margins.marginTop ?? 1440, unit, t, lang)}`,
    `${t('dialogs.pageSetup.bottom')} ${formatLayoutLength(margins.marginBottom ?? 1440, unit, t, lang)}`,
    `${t('dialogs.pageSetup.left')} ${formatLayoutLength(margins.marginLeft ?? 1440, unit, t, lang)}`,
    `${t('dialogs.pageSetup.right')} ${formatLayoutLength(margins.marginRight ?? 1440, unit, t, lang)}`,
  ].join('  ');
}

export function pageLayoutOrientation(properties?: SectionProperties): 'portrait' | 'landscape' {
  return properties?.orientation ?? ((properties?.pageWidth || 12240) > (properties?.pageHeight || 15840) ? 'landscape' : 'portrait');
}

export function pageOrientationPatch(properties: SectionProperties | undefined, orientation: 'portrait' | 'landscape'): Partial<SectionProperties> {
  if (orientation === pageLayoutOrientation(properties)) return { orientation };
  return { orientation, pageWidth: properties?.pageHeight || 15840, pageHeight: properties?.pageWidth || 12240 };
}
