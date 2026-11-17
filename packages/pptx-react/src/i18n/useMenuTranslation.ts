import { useMemo } from 'react';
import { createT, deepMerge } from '@betteroffice/pptx-i18n/editor';
import { menuEn } from '@betteroffice/pptx-i18n/menus';
import type { LocaleStrings, TFunction } from '@betteroffice/pptx-i18n';
import { useLocaleStrings } from './LocaleContext';

/** Translations that also cover the context menu messages, which load with the editing surface. */
export function useMenuTranslation(): { t: TFunction; lang: string } {
  const { strings, lang } = useLocaleStrings();
  const t = useMemo(
    () => createT(deepMerge(menuEn, strings as Record<string, unknown>) as LocaleStrings, lang),
    [strings, lang]
  );
  return { t, lang };
}
