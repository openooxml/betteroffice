import { describe, expect, test } from "bun:test";
import { isRtl, resolveLocale } from "./locale";
import en from "./i18n/en.json";
import ar from "./i18n/ar.json";
import de from "./i18n/de.json";
import es from "./i18n/es.json";
import fr from "./i18n/fr.json";
import he from "./i18n/he.json";
import hi from "./i18n/hi.json";
import id from "./i18n/id.json";
import it from "./i18n/it.json";
import ja from "./i18n/ja.json";
import ko from "./i18n/ko.json";
import nl from "./i18n/nl.json";
import pl from "./i18n/pl.json";
import ptBR from "./i18n/pt-BR.json";
import ru from "./i18n/ru.json";
import tr from "./i18n/tr.json";
import uk from "./i18n/uk.json";
import vi from "./i18n/vi.json";
import zhCN from "./i18n/zh-CN.json";
import zhTW from "./i18n/zh-TW.json";

describe("resolveLocale", () => {
  test("matches shipped codes exactly, ignoring case", () => {
    expect(resolveLocale(["de"])).toBe("de");
    expect(resolveLocale(["PT-BR"])).toBe("pt-BR");
    expect(resolveLocale(["Zh-Tw"])).toBe("zh-TW");
    expect(resolveLocale(["zh-cn"])).toBe("zh-CN");
  });

  test("walks languages in order", () => {
    expect(resolveLocale(["xx-YY", "fr-CA", "de"])).toBe("fr");
    expect(resolveLocale(["de-AT", "fr"])).toBe("de");
    expect(resolveLocale([])).toBe("en");
    expect(resolveLocale([""])).toBe("en");
    expect(resolveLocale(["xx", "yy-ZZ"])).toBe("en");
  });

  test("maps any Portuguese to pt-BR", () => {
    expect(resolveLocale(["pt"])).toBe("pt-BR");
    expect(resolveLocale(["pt-PT"])).toBe("pt-BR");
    expect(resolveLocale(["pt-AO"])).toBe("pt-BR");
  });

  test("maps Traditional Chinese tags to zh-TW", () => {
    expect(resolveLocale(["zh-Hant"])).toBe("zh-TW");
    expect(resolveLocale(["zh-Hant-TW"])).toBe("zh-TW");
    expect(resolveLocale(["zh-Hant-HK"])).toBe("zh-TW");
    expect(resolveLocale(["zh-TW"])).toBe("zh-TW");
    expect(resolveLocale(["zh-HK"])).toBe("zh-TW");
    expect(resolveLocale(["zh-MO"])).toBe("zh-TW");
    expect(resolveLocale(["zh-Hant-CN"])).toBe("zh-TW");
  });

  test("maps other Chinese tags to zh-CN", () => {
    expect(resolveLocale(["zh"])).toBe("zh-CN");
    expect(resolveLocale(["zh-Hans"])).toBe("zh-CN");
    expect(resolveLocale(["zh-CN"])).toBe("zh-CN");
    expect(resolveLocale(["zh-SG"])).toBe("zh-CN");
    expect(resolveLocale(["zh-Hans-HK"])).toBe("zh-CN");
  });

  test("falls back to the primary subtag when shipped", () => {
    expect(resolveLocale(["de-AT"])).toBe("de");
    expect(resolveLocale(["es-419"])).toBe("es");
    expect(resolveLocale(["en-GB"])).toBe("en");
    expect(resolveLocale(["fr-BE"])).toBe("fr");
    expect(resolveLocale(["uk-UA"])).toBe("uk");
  });

  test("maps legacy primary subtags", () => {
    expect(resolveLocale(["he-IL"])).toBe("he");
    expect(resolveLocale(["iw"])).toBe("he");
    expect(resolveLocale(["iw-IL"])).toBe("he");
    expect(resolveLocale(["in"])).toBe("id");
    expect(resolveLocale(["in-ID"])).toBe("id");
    expect(resolveLocale(["id-ID"])).toBe("id");
  });
});

describe("isRtl", () => {
  test("is true only for Arabic and Hebrew", () => {
    expect(isRtl("ar")).toBe(true);
    expect(isRtl("he")).toBe(true);
    expect(isRtl("en")).toBe(false);
    expect(isRtl("de")).toBe(false);
    expect(isRtl("zh-CN")).toBe(false);
  });
});

type Tree = { [key: string]: string | Tree };

function flatten(tree: Tree, prefix = ""): Map<string, string> {
  const out = new Map<string, string>();
  for (const [key, value] of Object.entries(tree)) {
    const path = prefix ? `${prefix}.${key}` : key;
    if (typeof value === "string") out.set(path, value);
    else for (const [k, v] of flatten(value, path)) out.set(k, v);
  }
  return out;
}

function placeholders(value: string): string[] {
  return [...value.matchAll(/\{(\w+)\}/g)].map((match) => match[1]).sort();
}

describe("desktop locale files", () => {
  const locales: Record<string, Tree> = {
    en,
    ar,
    de,
    es,
    fr,
    he,
    hi,
    id,
    it,
    ja,
    ko,
    nl,
    pl,
    "pt-BR": ptBR,
    ru,
    tr,
    uk,
    vi,
    "zh-CN": zhCN,
    "zh-TW": zhTW,
  };

  test("every locale has exactly en's keys and placeholders", () => {
    const expected = flatten(en as Tree);
    for (const [code, locale] of Object.entries(locales)) {
      const actual = flatten(locale);
      expect([...actual.keys()].sort(), code).toEqual([...expected.keys()].sort());
      for (const [key, value] of expected) {
        expect(placeholders(actual.get(key) ?? ""), `${code}:${key}`).toEqual(
          placeholders(value)
        );
      }
    }
  });
});
