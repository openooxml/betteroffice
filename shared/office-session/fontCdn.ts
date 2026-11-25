import {
  getFontCdnConfiguration as getRuntimeConfiguration,
  setFontCdnConfiguration as setRuntimeConfiguration,
} from './fontCdnRuntime';

/** @internal */
export interface FontCdnConfiguration {
  baseUrl: string | false;
  timeoutMs: number;
}

/** @internal */
export function getFontCdnConfiguration(): FontCdnConfiguration | undefined {
  return getRuntimeConfiguration();
}

/** @internal */
export function setFontCdnConfiguration(configuration: FontCdnConfiguration | undefined): void {
  setRuntimeConfiguration(configuration);
}
