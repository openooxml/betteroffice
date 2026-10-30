import {
  resolveBundledFamilyFace,
  resolveMetricCompatFace,
  resolveScriptFallbackFace,
  resolveLastResortChoice,
  resolveGlyphFallbackFace,
  resolveLightVariantFace,
  type BundledFontFace,
  type LastResortOffice,
  type BundledFontScript,
} from './manifest';

export type BundledFontLoader = (() => Promise<ArrayBuffer>) & {
  /** Word measurement family; defaults to the requested family. */
  measureAs?: string;
  onlyIfPrimaryFails?: boolean;
};

/** Structural provider contract keeps this package independent of the engine. */
export interface BundledFontSource {
  resolve(
    family: string,
    bold: boolean,
    italic: boolean,
  ): BundledFontLoader | undefined;
  /** Like `resolve`, but also accepts bundled family names such as `"Gelasio"`, and leaves missing styles for browser synthesis. */
  resolveFamily(
    family: string,
    bold: boolean,
    italic: boolean,
  ): (() => Promise<ArrayBuffer>) | undefined;
  resolveScriptFallback(
    script: BundledFontScript,
    bold: boolean,
    italic: boolean,
  ): (() => Promise<ArrayBuffer>) | undefined;
  /** Coverage for characters absent from a bundled substitute. */
  resolveGlyphFallback?(
    family: string,
    bold: boolean,
    italic: boolean,
    office?: LastResortOffice,
  ): (() => Promise<ArrayBuffer>) | undefined;
  resolveLastResort(
    family: string,
    bold: boolean,
    italic: boolean,
    office?: LastResortOffice,
  ): BundledFontLoader;
  /** A lighter member of a bundled family, such as `Calibri Light`, drawn with its regular face. */
  resolveLightVariant?(
    family: string,
    bold: boolean,
    italic: boolean,
  ): (() => Promise<ArrayBuffer>) | undefined;
}

export function fontProvider(
  loadBytes: (face: BundledFontFace) => Promise<ArrayBuffer>,
): BundledFontSource {
  const loaders = new Map<string, BundledFontLoader>();
  const load = (
    face: BundledFontFace,
    options?: Pick<BundledFontLoader, 'measureAs' | 'onlyIfPrimaryFails'>,
  ): BundledFontLoader => {
    const key = `${face.file}|${options?.measureAs ?? ''}|${options?.onlyIfPrimaryFails ?? false}`;
    let loader = loaders.get(key);
    if (loader === undefined) {
      loader = Object.assign(() => loadBytes(face), options);
      loaders.set(key, loader);
    }
    return loader;
  };
  return {
    resolve(family, bold, italic) {
      const face = resolveMetricCompatFace(family, bold, italic);
      return face ? load(face, { measureAs: face.family === 'Selawik' ? face.metricCompatWith : undefined }) : undefined;
    },
    resolveFamily(family, bold, italic) {
      const face = resolveBundledFamilyFace(family, bold, italic);
      return face ? load(face) : undefined;
    },
    resolveScriptFallback(script, bold, italic) {
      const face = resolveScriptFallbackFace(script, bold, italic);
      return face ? load(face) : undefined;
    },
    resolveGlyphFallback(family, bold, italic, office = 'powerpoint') {
      const face = resolveGlyphFallbackFace(family, bold, italic, office);
      return face ? load(face) : undefined;
    },
    resolveLightVariant(family, bold, italic) {
      const face = resolveLightVariantFace(family, bold, italic);
      return face ? load(face) : undefined;
    },
    resolveLastResort(family, bold, italic, office) {
      const { face, onlyIfPrimaryFails } = resolveLastResortChoice(family, bold, italic, office);
      return load(face, onlyIfPrimaryFails ? { onlyIfPrimaryFails } : undefined);
    },
  };
}
