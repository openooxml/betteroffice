import { createContext, useContext, useMemo } from 'react';
import type { ReactNode } from 'react';
import { createT, deepMerge, editorEn } from '@betteroffice/pptx-i18n/editor';
import type { LocaleStrings, TFunction, Translations } from '@betteroffice/pptx-i18n';

const LocaleContext = createContext<Translations>(editorEn);
const LangContext = createContext<string>('en');

export interface LocaleProviderProps {
  i18n?: Translations;
  children: ReactNode;
}

function formattingLocale(lang: unknown): string {
  if (typeof lang !== 'string') return 'en';
  try {
    return Intl.getCanonicalLocales(lang)[0] ?? 'en';
  } catch {
    return 'en';
  }
}

export function LocaleProvider({ i18n, children }: LocaleProviderProps) {
  const lang = formattingLocale(i18n?._lang);
  const merged = useMemo(
    () => deepMerge(editorEn as Record<string, unknown>, i18n as Record<string, unknown> | undefined),
    [i18n]
  );
  return (
    <LangContext.Provider value={lang}>
      <LocaleContext.Provider value={merged as Translations}>{children}</LocaleContext.Provider>
    </LangContext.Provider>
  );
}

export function useLocaleStrings(): { strings: Translations; lang: string } {
  const strings = useContext(LocaleContext);
  const lang = useContext(LangContext);
  return { strings, lang };
}

export function useTranslation(): { t: TFunction; lang: string } {
  const { strings, lang } = useLocaleStrings();
  const t = useMemo(() => createT(strings as LocaleStrings, lang), [strings, lang]);
  return { t, lang };
}
