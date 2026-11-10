import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const FONT_EXPORTS = 'maxDigitWidth, resolveGlyphFallbackFace, resolveLastResortFace, resolveMetricCompatFace, resolveScriptFallbackFace';
const checkoutFonts = resolve(dirname(fileURLToPath(import.meta.url)), '../../packages/fonts/src');

export function fontsModuleSource(publishedFonts) {
  return publishedFonts
    ? `export { createFontProvider } from '@betteroffice/fonts/cdn';\nexport { ${FONT_EXPORTS} } from '@betteroffice/fonts';\n`
    : `export { createFontProvider } from ${JSON.stringify(resolve(checkoutFonts, 'cdn.ts'))};\nexport { ${FONT_EXPORTS} } from ${JSON.stringify(resolve(checkoutFonts, 'manifest.ts'))};\n`;
}

export function qualityRendererPlugin(format, { publishedFonts = false } = {}) {
  if (!['pptx', 'xlsx', 'vsdx'].includes(format))
    throw new Error('Invalid Office capture format');
  const modules = {
    'virtual:office-quality-renderer': `export * from '@betteroffice/${format}';`,
    'virtual:office-quality-fonts': fontsModuleSource(publishedFonts),
  };
  return {
    name: 'office-quality-renderer',
    resolveId(source) {
      if (source in modules) return `\0${source}`;
    },
    load(source) {
      if (source.startsWith('\0') && source.slice(1) in modules) return modules[source.slice(1)];
    },
  };
}
