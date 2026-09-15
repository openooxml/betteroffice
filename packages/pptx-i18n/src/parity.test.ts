/** Every locale exposes exactly the key set of en.json. */
import { expect, test } from 'bun:test';
import de from '../de.json';
import en from '../en.json';
import fr from '../fr.json';
import he from '../he.json';
import hi from '../hi.json';
import id from '../id.json';
import pl from '../pl.json';
import ptBR from '../pt-BR.json';
import tr from '../tr.json';
import zhCN from '../zh-CN.json';

function leafKeys(value: unknown, prefix = ''): string[] {
  if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
    return Object.entries(value as Record<string, unknown>).flatMap(([key, child]) =>
      key === '_lang' && prefix === '' ? [] : leafKeys(child, prefix ? `${prefix}.${key}` : key),
    );
  }
  return [prefix];
}

const locales: Record<string, unknown> = {
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

const expected = new Set(leafKeys(en));

for (const [code, data] of Object.entries(locales)) {
  test(`${code} matches the en key set`, () => {
    const actual = new Set(leafKeys(data));
    const missing = [...expected].filter((key) => !actual.has(key));
    const extra = [...actual].filter((key) => !expected.has(key));
    expect({ missing, extra }).toEqual({ missing: [], extra: [] });
  });
}
