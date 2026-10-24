import { createRustMeasureSource, type ResidentFontRequirement } from '@betteroffice/docx/layout';
import type { EmbeddedFontFace, FontDefinition } from '@betteroffice/docx/utils';
import type { DocxFontEnvironment, YrsResidentFontRegistration } from '@betteroffice/docx/yrs';
import type { DocxEditorProps } from '../../DocxEditor';

export async function customWorkerFontFaces(fonts: ReadonlyArray<FontDefinition> | undefined, signal: AbortSignal): Promise<EmbeddedFontFace[]> {
  return Promise.all((fonts ?? []).map(async font => {
    const response = await fetch(font.src, { signal });
    if (!response.ok) throw new Error(`DOCX font ${font.family} failed to load: ${response.status}`);
    const data = await response.arrayBuffer();
    signal.throwIfAborted();
    const weight = font.weight === 'bold' || font.weight === 'bolder' || Number(font.weight) >= 600 ? 'bold' : 'normal';
    return { family: font.family, data, weight, style: 'normal', subsetted: false };
  }));
}

export function fontEnvironmentForRegistrations(fonts: YrsResidentFontRegistration[], chains: Record<string, number[]>,
  requirements: ResidentFontRequirement[], families: string[]): DocxFontEnvironment {
  const styles = new Map<number, ResidentFontRequirement>();
  for (const [key, slots] of Object.entries(chains)) {
    const [family, bold, italic] = key.split('|');
    const requirement = requirements.find(item => item.key === key) ?? { key, family: family!, bold: bold === '1', italic: italic === '1' };
    for (const slot of slots) if (!styles.has(slot)) styles.set(slot, requirement);
  }
  return { faces: fonts.map((font, index) => {
    if (!(font instanceof Uint8Array)) return { ...font };
    const style = styles.get(index);
    return { family: style?.family ?? 'Liberation Sans', data: font.slice(),
      weight: style?.bold ? 'bold' : 'normal', style: style?.italic ? 'italic' : 'normal' };
  }), chains: structuredClone(chains), families: [...new Set(families)] };
}

export async function prepareWorkerFontEnvironment(options: {
  fonts: ReadonlyArray<FontDefinition> | undefined;
  embedded: EmbeddedFontFace[];
  requirements: ResidentFontRequirement[];
  families: string[];
  provider: DocxEditorProps['measurementFontProvider'];
  compat: Parameters<ReturnType<typeof createRustMeasureSource>['setCompat']>[0];
  signal: AbortSignal;
}): Promise<DocxFontEnvironment> {
  const fonts: YrsResidentFontRegistration[] = [];
  const measurement = createRustMeasureSource({ bundled: options.provider, engine: {
    registerFont(bytes) { options.signal.throwIfAborted(); return fonts.push(bytes.slice()) - 1; },
    registerSubstituteFont(substituteOf, family) { return fonts.push({ substituteOf, family }) - 1; },
    clearFonts() { fonts.length = 0; },
  } });
  try {
    measurement.setCompat(options.compat);
    const custom = await customWorkerFontFaces(options.fonts, options.signal);
    measurement.setEmbeddedFaces([...custom, ...options.embedded]);
    const requirements = [...options.requirements, ...custom.map(face => ({
      key: `${face.family.toLowerCase()}|${face.weight === 'bold' ? 1 : 0}|0`, family: face.family,
      bold: face.weight === 'bold', italic: false,
    }))];
    await measurement.prepareFontRequirements(requirements);
    options.signal.throwIfAborted();
    const config = measurement.measurementConfigForRequirements(requirements);
    if (!config) throw new Error('DOCX worker fonts are not ready');
    return fontEnvironmentForRegistrations(fonts, config.fontChains, requirements,
      [...options.families, ...custom.map(face => face.family)]);
  } finally { measurement.setEmbeddedFaces([]); measurement.clear(); }
}
