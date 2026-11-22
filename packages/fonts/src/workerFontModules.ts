import * as fonts from './worker';
import * as progressive from './progressive-entry';
import { workerFontAssets } from './workerAssets';

async function requireFontPackage(specifier: string): Promise<void> {
  const base = await workerFontAssets('base');
  if (!base?.urls) throw new TypeError(`Failed to resolve module specifier '${specifier}'`);
}

export async function importFonts(): Promise<typeof fonts> {
  await requireFontPackage('@betteroffice/fonts');
  return fonts;
}

export async function importProgressiveFonts(): Promise<typeof progressive> {
  await requireFontPackage('@betteroffice/fonts/progressive');
  return progressive;
}
