import enJson from '../en.json';

export type LocaleStrings = typeof en;
export type LocaleCode = 'en';
export const en = enJson;
export const locales: Record<LocaleCode, PartialLocaleStrings> = { en };

export type DeepPartial<T> = { [K in keyof T]?: T[K] extends object ? DeepPartial<T[K]> : T[K] | null };
export type PartialLocaleStrings = DeepPartial<LocaleStrings> & { _lang?: LocaleCode | (string & {}) };
export type Translations = PartialLocaleStrings;
type DotPath<T, Prefix extends string = ''> = { [K in keyof T & string]: T[K] extends Record<string, unknown> ? DotPath<T[K], `${Prefix}${K}.`> : `${Prefix}${K}` }[keyof T & string];
export type TranslationKey = DotPath<LocaleStrings>;
type AnyRecord = Record<string, unknown>;
function isRecord(value: unknown): value is AnyRecord { return value !== null && typeof value === 'object' && !Array.isArray(value); }
export function deepMerge(base: AnyRecord, override: AnyRecord | undefined): AnyRecord { if (!override) return base; const result: AnyRecord = { ...base }; for (const key of Object.keys(override)) { const value = override[key]; if (value === null) continue; result[key] = isRecord(base[key]) && isRecord(value) ? deepMerge(base[key], value) : value ?? base[key]; } return result; }
function lookupKey(value: AnyRecord, path: string): string | undefined { let current: unknown = value; for (const part of path.split('.')) { if (!isRecord(current)) return undefined; current = current[part]; } return typeof current === 'string' ? current : undefined; }
export type TFunction = (key: TranslationKey, vars?: Record<string, string | number>) => string;
export function createT(strings: LocaleStrings, _lang = 'en'): TFunction { return (key, vars) => (lookupKey(strings as AnyRecord, key) ?? key).replace(/\{(\w+)\}/g, (_, name) => vars?.[name] === undefined ? `{${name}}` : String(vars[name])); }
export type DiagnosticCategory = 'integrity' | 'fidelity';
export function diagnosticMessage(t: TFunction, category: DiagnosticCategory, code: string): string { const key = `diagnostics.${category}.${code}` as TranslationKey; const message = t(key); return message === key ? t('diagnostics.unknown', { category, code }) : message; }
