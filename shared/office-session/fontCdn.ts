import { workerFontAssets } from './optionalFontPacks';

export interface FontCdnConfiguration {
  baseUrl: string | false;
  timeoutMs: number;
  implicit?: true;
}

type FontGlobals = {
  window?: { document?: unknown };
  process?: { versions?: { node?: string } };
  Deno?: unknown;
};

export const JSDELIVR = 'https://cdn.jsdelivr.net/npm/';
export const STALL_MS = 10_000;
const CONFIGURATION = Symbol.for('betteroffice.fonts.cdn');
const CONFIGURED_HERE = Symbol.for('betteroffice.fonts.cdn.configured-here');

const scope = globalThis as FontGlobals & {
  [CONFIGURATION]?: FontCdnConfiguration;
  [CONFIGURED_HERE]?: true;
};

function isServer(globals: FontGlobals): boolean {
  return globals.process?.versions?.node !== undefined || globals.Deno !== undefined;
}

export function isBrowserFontPage(globals: FontGlobals): boolean {
  return globals.window?.document != null && !isServer(globals);
}

function configuredFontCdn(configuration: FontCdnConfiguration | undefined): FontCdnConfiguration | undefined {
  if (configuration && !(configuration.implicit && isServer(scope))) return { ...configuration };
}

export function getFontCdnConfiguration(): FontCdnConfiguration | undefined {
  const configured = scope[CONFIGURATION];
  if (configured) return configuredFontCdn(configured);
  if (isBrowserFontPage(scope)) return { baseUrl: JSDELIVR, timeoutMs: STALL_MS, implicit: true };
}

export function setFontCdnConfiguration(configuration: FontCdnConfiguration | undefined, configuredHere?: true): void {
  if (configuration === undefined) delete scope[CONFIGURATION];
  else scope[CONFIGURATION] = { ...configuration };
  if (configuredHere) scope[CONFIGURED_HERE] = configuredHere;
  else delete scope[CONFIGURED_HERE];
}

export function workerFontCdnConfiguration(reply: { cdn?: FontCdnConfiguration } | undefined): FontCdnConfiguration | undefined {
  return configuredFontCdn(scope[CONFIGURED_HERE] || reply === undefined ? scope[CONFIGURATION] : reply.cdn);
}

export function fontCdnCacheKey(configuration: FontCdnConfiguration | null = getFontCdnConfiguration() ?? null): string {
  return configuration ? JSON.stringify([configuration.baseUrl, configuration.timeoutMs, configuration.implicit]) : '';
}

export async function pngFontCdnCacheKey(): Promise<string> {
  if (scope[CONFIGURED_HERE] || isServer(scope) || isBrowserFontPage(scope)) return fontCdnCacheKey();
  const reply = await workerFontAssets('base').catch(() => undefined);
  return fontCdnCacheKey(workerFontCdnConfiguration(reply) ?? null);
}
