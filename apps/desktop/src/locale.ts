import type { LocaleCode } from "@betteroffice/docx-i18n";

const shipped: readonly LocaleCode[] = [
  "en",
  "ar",
  "de",
  "es",
  "fr",
  "he",
  "hi",
  "id",
  "it",
  "ja",
  "ko",
  "nl",
  "pl",
  "pt-BR",
  "ru",
  "tr",
  "uk",
  "vi",
  "zh-CN",
  "zh-TW",
];

const byLowerCase = new Map(shipped.map((code) => [code.toLowerCase(), code]));

const primaryCodes = new Map<string, LocaleCode>([
  ["en", "en"],
  ["ar", "ar"],
  ["de", "de"],
  ["es", "es"],
  ["fr", "fr"],
  ["he", "he"],
  ["iw", "he"],
  ["hi", "hi"],
  ["id", "id"],
  ["in", "id"],
  ["it", "it"],
  ["ja", "ja"],
  ["ko", "ko"],
  ["nl", "nl"],
  ["pl", "pl"],
  ["ru", "ru"],
  ["tr", "tr"],
  ["uk", "uk"],
  ["vi", "vi"],
]);

export function resolveLocale(languages: readonly string[]): LocaleCode {
  for (const language of languages) {
    const tag = language.trim().replace(/_/g, "-");
    if (!tag) continue;
    const exact = byLowerCase.get(tag.toLowerCase());
    if (exact) return exact;
    const parts = tag.toLowerCase().split("-").filter(Boolean);
    const primary = parts[0];
    if (!primary) continue;
    if (primary === "pt") return "pt-BR";
    if (primary === "zh") {
      const traditional =
        parts.includes("hant") ||
        parts.some((part) => ["tw", "hk", "mo"].includes(part));
      return traditional ? "zh-TW" : "zh-CN";
    }
    const byPrimary = primaryCodes.get(primary);
    if (byPrimary) return byPrimary;
  }
  return "en";
}

export function isRtl(code: string): boolean {
  const primary = code.toLowerCase().split("-")[0];
  return primary === "ar" || primary === "he" || primary === "iw";
}
