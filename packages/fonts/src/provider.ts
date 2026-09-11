import {
  resolveMetricCompatFace,
  resolveScriptFallbackFace,
  resolveLastResortFace,
  type BundledFontFace,
  type BundledFontScript,
} from './manifest';

/** Structural provider contract keeps this package independent of the engine. */
export interface BundledFontSource {
  resolve(
    family: string,
    bold: boolean,
    italic: boolean,
  ): (() => Promise<ArrayBuffer>) | undefined;
  resolveScriptFallback(
    script: BundledFontScript,
    bold: boolean,
    italic: boolean,
  ): (() => Promise<ArrayBuffer>) | undefined;
  resolveLastResort(
    family: string,
    bold: boolean,
    italic: boolean,
  ): () => Promise<ArrayBuffer>;
}

export function fontProvider(
  loadBytes: (face: BundledFontFace) => Promise<ArrayBuffer>,
): BundledFontSource {
  const load = (face: BundledFontFace) => () => loadBytes(face);
  return {
    resolve(family, bold, italic) {
      const face = resolveMetricCompatFace(family, bold, italic);
      return face ? load(face) : undefined;
    },
    resolveScriptFallback(script, bold, italic) {
      const face = resolveScriptFallbackFace(script, bold, italic);
      return face ? load(face) : undefined;
    },
    resolveLastResort(family, bold, italic) {
      return load(resolveLastResortFace(family, bold, italic));
    },
  };
}
