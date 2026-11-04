import { useMemo } from 'react';
import { createT, deepMerge, proposalEn } from '@betteroffice/pptx-i18n';
import type { LocaleStrings, TFunction } from '@betteroffice/pptx-i18n';
import { useLocaleStrings } from './LocaleContext';

export function useProposalTranslation(): { t: TFunction; lang: string } {
  const { strings, lang } = useLocaleStrings();
  const t = useMemo(
    () => createT(deepMerge(proposalEn, strings as Record<string, unknown>) as LocaleStrings, lang),
    [strings, lang]
  );
  return { t, lang };
}
