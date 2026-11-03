import { createRustMeasureSource, fontRequirementCharacters, type ResidentFontRequirement } from '@betteroffice/docx/layout';
import type { EmbeddedFontFace, FontDefinition } from '@betteroffice/docx/utils';
import type { DocxFontEnvironment, YrsResidentFontRegistration } from '@betteroffice/docx/yrs';
import type { DocxEditorProps } from '../../DocxEditor';
import { hostedFontProvider } from '../../../hooks/hostedFonts';

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
  requirements: ResidentFontRequirement[], families: string[],
  registeredStyles?: ReadonlyMap<number, Pick<ResidentFontRequirement, 'family' | 'bold' | 'italic'>>): DocxFontEnvironment {
  const styles = new Map<number, Pick<ResidentFontRequirement, 'family' | 'bold' | 'italic'>>(registeredStyles);
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
  }), chains: structuredClone(chains), families: [...new Set([...families, ...Array.from(registeredStyles?.values() ?? [], face => face.family)])] };
}

/**
 * `prepared` extended by every requirement of `next` it does not cover, or null when it covers all.
 * A requirement covers one with its key only with the same scripts in the same order: they fix the chain.
 * New characters alone count only when `covers` cannot place them in the fonts already registered;
 * it judges only the requirements that gained characters.
 */
export function uncoveredFontRequirements(prepared: readonly ResidentFontRequirement[],
  next: readonly ResidentFontRequirement[],
  covers?: (requirements: ResidentFontRequirement[]) => boolean): ResidentFontRequirement[] | null {
  const byKey = new Map(prepared.map(requirement => [requirement.key, requirement]));
  const added: ResidentFontRequirement[] = [];
  let chainsChanged = false;
  for (const requirement of next) {
    const known = byKey.get(requirement.key);
    const characters = fontRequirementCharacters(requirement);
    const knownCharacters = known ? fontRequirementCharacters(known) : [];
    const covered = new Set(knownCharacters);
    const sameChain = !!known && (known.scripts ?? []).join() === (requirement.scripts ?? []).join();
    if (sameChain && characters.every(character => covered.has(character))) continue;
    const extended = known ? {
      ...requirement, characters: [...new Set([...knownCharacters, ...characters])].sort(),
    } : requirement;
    byKey.set(requirement.key, extended);
    added.push(extended);
    chainsChanged ||= !sameChain;
  }
  if (added.length === 0) return null;
  return !chainsChanged && covers?.(added) ? null : [...byKey.values()];
}

export async function prepareWorkerFontEnvironment(options: {
  fonts: ReadonlyArray<FontDefinition> | undefined;
  embedded: EmbeddedFontFace[];
  requirements: ResidentFontRequirement[];
  families: string[];
  provider: DocxEditorProps['measurementFontProvider'];
  runtime?: Parameters<typeof createRustMeasureSource>[0]['runtime'];
  compat: Parameters<ReturnType<typeof createRustMeasureSource>['setCompat']>[0];
  signal: AbortSignal;
  onCoverage?: (covers: (requirements: ResidentFontRequirement[]) => boolean) => void;
}): Promise<DocxFontEnvironment> {
  const fonts: YrsResidentFontRegistration[] = [];
  const styles = new Map<number, Pick<ResidentFontRequirement, 'family' | 'bold' | 'italic'>>();
  const measurement = createRustMeasureSource({ bundled: hostedFontProvider(options.provider, options.fonts), runtime: options.runtime, engine: {
    registerFont(bytes, face) {
      options.signal.throwIfAborted();
      const id = fonts.push(bytes.slice()) - 1;
      if (face) styles.set(id, face);
      return id;
    },
    registerSubstituteFont(substituteOf, family) { return fonts.push({ substituteOf, family }) - 1; },
    clearFonts() { fonts.length = 0; },
  } });
  try {
    measurement.setCompat(options.compat);
    const custom = await customWorkerFontFaces(options.fonts, options.signal);
    measurement.setEmbeddedFaces(options.embedded, custom);
    const requirements = [...options.requirements, ...custom.map(face => ({
      key: `${face.family.toLowerCase()}|${face.weight === 'bold' ? 1 : 0}|0`, family: face.family,
      bold: face.weight === 'bold', italic: false,
    }))];
    await measurement.prepareFontRequirements(requirements);
    options.signal.throwIfAborted();
    const config = measurement.measurementConfigForRequirements(requirements);
    if (!config) throw new Error('DOCX worker fonts are not ready');
    options.onCoverage?.(measurement.coverageCheck(requirements));
    return fontEnvironmentForRegistrations(fonts, config.fontChains, requirements,
      [...options.families, ...custom.map(face => face.family)], styles);
  } finally { measurement.setEmbeddedFaces([]); measurement.clear(); }
}
