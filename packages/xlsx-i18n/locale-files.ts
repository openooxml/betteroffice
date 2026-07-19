import { readdirSync } from 'node:fs';

/** Matches BCP-47-shaped JSON filenames. */
export const BCP47_FILENAME = /^[a-z]{2,3}(-[a-zA-Z0-9]{2,8})*\.json$/;

/** Returns the sorted locale codes in an i18n package directory. */
export function readLocaleCodes(i18nDir: string): string[] {
  return readdirSync(i18nDir)
    .filter((file) => BCP47_FILENAME.test(file))
    .map((file) => file.replace(/\.json$/, ''))
    .sort();
}
