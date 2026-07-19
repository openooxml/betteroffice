/**
 * Shared locale data, types, and runtime helpers for @betteroffice/pptx-react.
 *
 * @packageDocumentation
 * @public
 */

import enJson from '../en.json';
import deJson from '../de.json';
import frJson from '../fr.json';
import heJson from '../he.json';
import hiJson from '../hi.json';
import idJson from '../id.json';
import plJson from '../pl.json';
import ptBRJson from '../pt-BR.json';
import trJson from '../tr.json';
import zhCNJson from '../zh-CN.json';

export type LocaleStrings = typeof enJson;
export type LocaleCode = 'en' | 'de' | 'fr' | 'he' | 'hi' | 'id' | 'pl' | 'pt-BR' | 'tr' | 'zh-CN';

export const en: LocaleStrings = enJson;
export const de: PartialLocaleStrings = deJson;
export const fr: PartialLocaleStrings = frJson;
export const he: PartialLocaleStrings = heJson;
export const hi: PartialLocaleStrings = hiJson;
export const id: PartialLocaleStrings = idJson;
export const pl: PartialLocaleStrings = plJson;
export const ptBR: PartialLocaleStrings = ptBRJson;
export const tr: PartialLocaleStrings = trJson;
export const zhCN: PartialLocaleStrings = zhCNJson;

export const locales: Record<LocaleCode, PartialLocaleStrings> = {
  en,
  de,
  fr,
  he,
  hi,
  id,
  pl,
  'pt-BR': ptBR,
  tr,
  'zh-CN': zhCN,
};

export type DeepPartial<T> = {
  [K in keyof T]?: T[K] extends object ? DeepPartial<T[K]> : T[K] | null;
};

export type PartialLocaleStrings = DeepPartial<LocaleStrings> & {
  _lang?: LocaleCode | (string & {});
};

export type Translations = PartialLocaleStrings;

type DotPath<T, Prefix extends string = ''> = {
  [K in keyof T & string]: T[K] extends Record<string, unknown>
    ? DotPath<T[K], `${Prefix}${K}.`>
    : `${Prefix}${K}`;
}[keyof T & string];

export type TranslationKey = DotPath<LocaleStrings>;

type AnyRecord = Record<string, unknown>;

function isRecord(v: unknown): v is AnyRecord {
  return v !== null && typeof v === 'object' && !Array.isArray(v);
}

export function deepMerge(base: AnyRecord, override: AnyRecord | undefined): AnyRecord {
  if (!override) return base;
  const result: AnyRecord = { ...base };
  for (const key of Object.keys(override)) {
    const overVal = override[key];
    if (overVal === null) continue;
    if (isRecord(base[key]) && isRecord(overVal)) {
      result[key] = deepMerge(base[key], overVal);
    } else if (overVal !== undefined) {
      result[key] = overVal;
    }
  }
  return result;
}

function lookupKey(obj: AnyRecord, path: string): string | undefined {
  let current: unknown = obj;
  for (const part of path.split('.')) {
    if (!isRecord(current)) return undefined;
    current = current[part];
  }
  return typeof current === 'string' ? current : undefined;
}

function parseBranches(branchStr: string): Record<string, string> {
  const parsed: Record<string, string> = {};
  const regex = /(=\d+|\w+)\s*\{([^}]*)\}/g;
  let match;
  while ((match = regex.exec(branchStr)) !== null) {
    parsed[match[1]] = match[2];
  }
  return parsed;
}

function formatMessage(
  template: string,
  vars?: Record<string, string | number>,
  lang?: string
): string {
  if (!vars) return template;

  const result = template.replace(
    /\{(\w+),\s*plural,((?:[^{}]|\{[^{}]*\})*)\}/g,
    (full, varName, branchStr) => {
      const count = Number(vars[varName]);
      if (isNaN(count)) return full;
      const parsed = parseBranches(branchStr);
      const exact = parsed[`=${count}`];
      if (exact !== undefined) return exact.replace(/#/g, String(count));
      let category: string;
      try {
        category = new Intl.PluralRules(lang || 'en').select(count);
      } catch {
        category = count === 1 ? 'one' : 'other';
      }
      const text = parsed[category] ?? parsed['other'] ?? '';
      return text.replace(/#/g, String(count));
    }
  );

  return result.replace(/\{(\w+)\}/g, (_, key) => {
    const val = vars[key];
    return val !== undefined ? String(val) : `{${key}}`;
  });
}

export type TFunction = (
  key: TranslationKey,
  vars?: Record<string, string | number>
) => string;

export function createT(strings: LocaleStrings, lang = 'en'): TFunction {
  return (key, vars) => {
    const value = lookupKey(strings as AnyRecord, key);
    return formatMessage(value ?? key, vars, lang);
  };
}
