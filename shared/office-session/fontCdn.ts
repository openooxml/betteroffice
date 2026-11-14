/** @internal */
export interface FontCdnConfiguration {
  baseUrl: string | false;
  timeoutMs: number;
}

const key = Symbol.for('betteroffice.fonts.cdn');
type FontCdnScope = { [key]?: FontCdnConfiguration };

/** @internal */
export function getFontCdnConfiguration(): FontCdnConfiguration | undefined {
  const configuration = (globalThis as FontCdnScope)[key];
  return configuration === undefined ? undefined : { ...configuration };
}

/** @internal */
export function setFontCdnConfiguration(configuration: FontCdnConfiguration | undefined): void {
  const scope = globalThis as FontCdnScope;
  if (configuration === undefined) delete scope[key];
  else scope[key] = { ...configuration };
}
