import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const FONT_EXPORTS = 'maxDigitWidth, resolveGlyphFallbackFace, resolveLastResortFace, resolveMetricCompatFace, resolveScriptFallbackFace';
const DOCX_FONT_EXPORTS = 'BUNDLED_FONTS, resolveGlyphFallbackFace, resolveLastResortChoice, resolveMetricCompatFace, resolveScriptFallbackFace';
const checkoutFonts = resolve(dirname(fileURLToPath(import.meta.url)), '../../packages/fonts/src');

export function fontsModuleSource(publishedFonts, names = FONT_EXPORTS) {
  return publishedFonts
    ? `export { createFontProvider } from '@betteroffice/fonts/cdn';\nexport { ${names} } from '@betteroffice/fonts';\n`
    : `export { createFontProvider } from ${JSON.stringify(resolve(checkoutFonts, 'cdn.ts'))};\nexport { ${names} } from ${JSON.stringify(resolve(checkoutFonts, 'manifest.ts'))};\n`;
}

function virtualModules(name, modules) {
  return {
    name,
    resolveId(source) {
      if (source in modules) return `\0${source}`;
    },
    load(source) {
      if (source.startsWith('\0') && source.slice(1) in modules) return modules[source.slice(1)];
    },
  };
}

export function qualityRendererPlugin(format, { publishedFonts = false } = {}) {
  if (!['pptx', 'xlsx', 'vsdx'].includes(format))
    throw new Error('Invalid Office capture format');
  return virtualModules('office-quality-renderer', {
    'virtual:office-quality-renderer': `export * from '@betteroffice/${format}';`,
    'virtual:office-quality-fonts': fontsModuleSource(publishedFonts),
  });
}

/** The DOCX harness's font provider and manifest, from the published release or this checkout. */
export function docxFontsPlugin({ publishedFonts = false } = {}) {
  const manifest = publishedFonts ? "'@betteroffice/fonts'" : JSON.stringify(resolve(checkoutFonts, 'manifest.ts'));
  return virtualModules('docx-quality-fonts', {
    'virtual:docx-quality-fonts': `${fontsModuleSource(publishedFonts, DOCX_FONT_EXPORTS)}import * as manifest from ${manifest};\nexport const BASE_LATIN_FONTS = manifest.BASE_LATIN_FONTS ?? [];\n`,
  });
}
