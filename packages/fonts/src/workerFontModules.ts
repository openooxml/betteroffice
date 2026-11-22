import { WorkerFontModuleUnavailableError } from '../../../shared/office-session/optionalFontPacks';
import * as fonts from './worker';
import * as progressive from './progressive-entry';
import { workerFontAssets } from './workerAssets';

async function requireFontPackage(specifier: string): Promise<void> {
  const base = await workerFontAssets('base');
  if (!base?.urls) throw new WorkerFontModuleUnavailableError(specifier);
}

export async function importFonts(): Promise<typeof fonts> {
  await requireFontPackage('@betteroffice/fonts');
  return fonts;
}

export async function importProgressiveFonts(): Promise<typeof progressive> {
  await requireFontPackage('@betteroffice/fonts/progressive');
  return progressive;
}
