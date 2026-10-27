import { createT, deepMerge, type LocaleCode, type LocaleStrings } from "@betteroffice/docx-i18n";
import enJson from "./en.json";
import arJson from "./ar.json";
import deJson from "./de.json";
import esJson from "./es.json";
import frJson from "./fr.json";
import heJson from "./he.json";
import hiJson from "./hi.json";
import idJson from "./id.json";
import itJson from "./it.json";
import jaJson from "./ja.json";
import koJson from "./ko.json";
import nlJson from "./nl.json";
import plJson from "./pl.json";
import ptBRJson from "./pt-BR.json";
import ruJson from "./ru.json";
import trJson from "./tr.json";
import ukJson from "./uk.json";
import viJson from "./vi.json";
import zhCNJson from "./zh-CN.json";
import zhTWJson from "./zh-TW.json";
import { isRtl, resolveLocale } from "../locale";

export type DesktopStrings = typeof enJson;

type DotPath<T, Prefix extends string = ""> = {
  [K in keyof T & string]: T[K] extends Record<string, unknown>
    ? DotPath<T[K], `${Prefix}${K}.`>
    : `${Prefix}${K}`;
}[keyof T & string];

export type DesktopKey = DotPath<DesktopStrings>;

export type DesktopT = (
  key: DesktopKey,
  vars?: Record<string, string | number>
) => string;

const strings: Record<LocaleCode, DesktopStrings> = {
  en: enJson,
  ar: arJson as DesktopStrings,
  de: deJson as DesktopStrings,
  es: esJson as DesktopStrings,
  fr: frJson as DesktopStrings,
  he: heJson as DesktopStrings,
  hi: hiJson as DesktopStrings,
  id: idJson as DesktopStrings,
  it: itJson as DesktopStrings,
  ja: jaJson as DesktopStrings,
  ko: koJson as DesktopStrings,
  nl: nlJson as DesktopStrings,
  pl: plJson as DesktopStrings,
  "pt-BR": ptBRJson as DesktopStrings,
  ru: ruJson as DesktopStrings,
  tr: trJson as DesktopStrings,
  uk: ukJson as DesktopStrings,
  vi: viJson as DesktopStrings,
  "zh-CN": zhCNJson as DesktopStrings,
  "zh-TW": zhTWJson as DesktopStrings,
};

export const locale: LocaleCode = resolveLocale(
  typeof navigator === "undefined"
    ? []
    : navigator.languages?.length
      ? navigator.languages
      : [navigator.language]
);

if (typeof document !== "undefined") {
  document.documentElement.lang = locale;
  document.documentElement.dir = isRtl(locale) ? "rtl" : "ltr";
}

const merged = deepMerge(
  enJson as unknown as Record<string, unknown>,
  strings[locale] as unknown as Record<string, unknown>
) as unknown as LocaleStrings;

export const t: DesktopT = createT(merged, locale) as unknown as DesktopT;
