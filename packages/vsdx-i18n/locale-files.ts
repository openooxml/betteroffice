import { readdirSync } from 'node:fs';

export const BCP47_FILENAME = /^[a-z]{2,3}(-[a-zA-Z0-9]{2,8})*\.json$/;

export function readLocaleCodes(i18nDir: string): string[] {
  return readdirSync(i18nDir).filter((file) => BCP47_FILENAME.test(file)).map((file) => file.replace(/\.json$/, '')).sort();
}
