import type { ProgressiveFontSpec } from '../../yrs/residentProgressiveFonts';
import type { CompatibilityFlags } from '../../docx/settingsParser';
import type { FontDecoder } from '../../../../../shared/office-session/fontDecoder';
import { baseLatinAlias, baseLatinCandidate, charactersFromBits, symbolFamily, type FontCoverage } from './fontCoverage';
import {
  excludeHostedFontFamilies,
  hostedFontFamilies,
  IMPLICIT_FONT_COVERAGE,
  isImplicitFontProvider,
  resolveDefaultFontProvider,
  baseLatinFontState,
} from './defaultFontProvider';
import {
  TextMeasureFontRegistry,
  type BundledFontProvider,
  type EmbeddedFaceInput,
  type FontScript,
} from './fontRegistry';

export interface RustTextEngine {
  registerFont(bytes: Uint8Array, face?: { family: string; bold: boolean; italic: boolean }): number;
  /**
   * Measurement view of `id` carrying the vertical metrics and advance pitch
   * Word measures `family` with; `id` when the engine knows none. Optional so
   * a host that predates it still measures, with the substitute's own
   * metrics.
   */
  registerSubstituteFont?(id: number, family: string): number;
  clearFonts(): void;
}

export interface ResidentFontRequirement {
  key: string;
  family: string;
  bold: boolean;
  italic: boolean;
  scripts?: FontScript[];
  characters?: string[];
  characterBits?: string;
}

/** @internal */
export function fontRequirementCharacters(requirement: ResidentFontRequirement): string[] {
  return [...new Set([...(requirement.characters ?? []), ...charactersFromBits(requirement.characterBits)])];
}

function mergeRequirements(requirements: ResidentFontRequirement[]): ResidentFontRequirement[] {
  const merged = new Map<string, ResidentFontRequirement>();
  for (const requirement of requirements) {
    const characters = fontRequirementCharacters(requirement);
    const known = merged.get(requirement.key);
    merged.set(requirement.key, known ? {
      ...known,
      scripts: [...new Set([...(known.scripts ?? []), ...(requirement.scripts ?? [])])],
      characters: [...new Set([...(known.characters ?? []), ...characters])],
    } : { ...requirement, characters });
  }
  return [...merged.values()];
}

export interface ResidentMeasurementConfig {
  fontChains: Record<string, number[]>;
  defaults: { fontSize: number; fontFamily: string };
  compat: {
    compatibilityMode?: number;
    noLeading: boolean;
    doNotExpandShiftReturn: boolean;
  };
  authoritativeShaping: true;
}

export interface RustMeasureSource {
  setEmbeddedFaces(faces: EmbeddedFaceInput[], appFaces?: EmbeddedFaceInput[]): void;
  setCompat(flags: CompatibilityFlags | undefined): void;
  prepareFontRequirements(
    requirements: ResidentFontRequirement[],
  ): Promise<boolean>;
  measurementConfigForRequirements(
    requirements: ResidentFontRequirement[],
  ): ResidentMeasurementConfig | undefined;
  /** @internal Whether later requirements need no face the prepared chains lack (characters the base Latin tier settled count as covered); holds no font bytes, so it outlives `clear()`. */
  coverageCheck(
    requirements: ResidentFontRequirement[],
  ): (next: ResidentFontRequirement[]) => boolean;
  clear(): void;
  /** The script-fallback slots registered from a sliced face's skeleton and core, with their files. */
  progressiveFonts?(): Promise<ProgressiveFontSpec[]>;
}

let enginePromise: Promise<RustTextEngine> | null = null;

export function getRustTextEngine(): Promise<RustTextEngine> {
  enginePromise ??= import('../wasm/index').then(async (module) => {
    await module.preloadLayoutWasm();
    return {
      registerFont: module.registerMeasureFont,
      registerSubstituteFont: module.registerSubstituteMeasureFont,
      clearFonts: module.clearMeasureFonts,
    };
  });
  return enginePromise;
}

/** Builds a source with an explicit or default font provider. */
export function createRustMeasureSource(options: {
  engine: RustTextEngine;
  bundled?: BundledFontProvider;
  runtime?: { decoder?: FontDecoder };
  /** Register sliced script-fallback faces from their skeleton and core. */
  progressive?: boolean;
  /** With the default provider, give families it doesn't carry the last-resort face. */
  standIns?: boolean;
}): RustMeasureSource {
  const engine = options.engine;
  const hosted = hostedFontFamilies(options.bundled);
  let implicit = false;
  let generation = 0;
  let appFamilies = new Set<string>();
  const hostedCoverageUnknown = (family: string) => implicit &&
    hosted?.has(family.trim().toLowerCase()) === true && !appFamilies.has(family.trim().toLowerCase());
  const registry = new TextMeasureFontRegistry(
    {
      registerFont: (bytes, face) => face ? engine.registerFont(bytes, face) : engine.registerFont(bytes),
      ...(engine.registerSubstituteFont && {
        registerSubstituteFont: (id, family) =>
          engine.registerSubstituteFont!(id, family),
      }),
    },
    { baseLatin: () => import('./baseLatinFaces').then(tier => tier.resolveBaseLatinProvider(baseLatinFontState(), options.runtime)), bundled: (hosted ? undefined : options.bundled) ?? (async () => {
      const current = generation;
      const provider = excludeHostedFontFamilies(await resolveDefaultFontProvider(options.runtime, options.standIns), hosted);
      if (current === generation) implicit = isImplicitFontProvider(provider);
      return provider;
    }), progressive: options.progressive
      ? async (slices) => import('@betteroffice/fonts/progressive').then(async ({ openSlicedFace, sliceSourceOf }) => {
        const face = await openSlicedFace(slices as Parameters<typeof openSlicedFace>[0]);
        if (!sliceSourceOf(face)) throw new Error('A sliced face opened without file URLs cannot load its outlines in the worker');
        return face;
      })
      : undefined },
  );
  let compat: CompatibilityFlags | undefined;

  return {
    async progressiveFonts() {
      const specs: ProgressiveFontSpec[] = [];
      if (registry.progressiveFaces().size === 0) return specs;
      const { sliceSourceOf } = await import('@betteroffice/fonts/progressive');
      for (const [slot, face] of registry.progressiveFaces()) {
        const source = await sliceSourceOf(face as Parameters<typeof sliceSourceOf>[0]);
        if (source) specs.push({ slot, source });
      }
      return specs;
    },

    setEmbeddedFaces(faces, appFaces): void {
      appFamilies = new Set((appFaces ?? []).map(face => face.family.trim().toLowerCase()));
      registry.setEmbeddedFaces(faces, appFaces);
    },

    setCompat(flags): void {
      compat = flags;
    },

    async prepareFontRequirements(requirements): Promise<boolean> {
      requirements = mergeRequirements(requirements);
      const settled = await Promise.all(
        requirements.flatMap((requirement) => [
          (async () => {
            try {
              const cached = registry.getCachedFontIdChain(
                requirement.family,
                requirement.bold,
                requirement.italic,
                false,
              );
              if (cached !== undefined) return false;
              await registry.getFontIdChain(
                requirement.family,
                requirement.bold,
                requirement.italic,
              );
              return true;
            } catch {
              return false;
            }
          })(),
          ...Array.from(new Set(requirement.scripts ?? []), async (script) => {
            try {
              if (hosted?.has(requirement.family.trim().toLowerCase())) {
                await registry.getFontIdChain(requirement.family, requirement.bold, requirement.italic);
                if (implicit) return false;
              }
              if (
                registry.getCachedScriptFallbackIds([script], false) !==
                undefined
              )
                return false;
              await registry.getScriptFallbackIds([script]);
              return true;
            } catch {
              return false;
            }
          }),
        ]),
      );
      const backstops = await Promise.all(requirements.map(async (requirement) => {
        if (hostedCoverageUnknown(requirement.family)) return false;
        const earlier = [
          ...(registry.getCachedFontIdChain(requirement.family, requirement.bold, requirement.italic, false) ?? []),
          ...(implicit && hosted?.has(requirement.family.trim().toLowerCase()) ? [] :
            registry.getCachedScriptFallbackIds(requirement.scripts ?? [], false) ?? []),
        ];
        return registry.prepareBaseLatin(requirement.family, requirement.bold, requirement.italic,
          requirement.characters ?? [], earlier, requirement.key);
      }));
      return settled.some(Boolean) || backstops.some(Boolean);
    },

    measurementConfigForRequirements(
      requirements,
    ): ResidentMeasurementConfig | undefined {
      requirements = mergeRequirements(requirements);
      const settled = requirements.map((requirement) => ({
        requirement,
        familyIds: registry.getCachedFontIdChain(
          requirement.family,
          requirement.bold,
          requirement.italic,
          false,
        ),
        scriptIds: implicit && hosted?.has(requirement.family.trim().toLowerCase()) ? [] : registry.getCachedScriptFallbackIds(
          requirement.scripts ?? [],
          false,
        ),
      }));
      if (
        settled.some(
          ({ requirement, familyIds, scriptIds }) =>
            familyIds === undefined || scriptIds === undefined ||
            (!hostedCoverageUnknown(requirement.family) &&
              registry.getCachedBaseLatinIds(requirement.family, requirement.bold, requirement.italic,
                requirement.characters, [...familyIds, ...scriptIds], requirement.key) === undefined),
        )
      )
        return undefined;
      const fontChains: Record<string, number[]> = {};
      for (const { requirement } of settled) {
        const familyIds = registry.getCachedFontIdChain(
          requirement.family,
          requirement.bold,
          requirement.italic,
        );
        const scriptIds = implicit && hosted?.has(requirement.family.trim().toLowerCase()) ? [] : registry.getCachedScriptFallbackIds(
          requirement.scripts ?? [],
        );
        if (familyIds === undefined || scriptIds === undefined)
          return undefined;
        const backstopIds = hostedCoverageUnknown(requirement.family) ? [] :
          registry.getCachedBaseLatinIds(requirement.family, requirement.bold, requirement.italic,
            requirement.characters, [...familyIds, ...scriptIds], requirement.key);
        if (backstopIds === undefined) return undefined;
        if (implicit && familyIds.length === 0 && backstopIds.length === 0) continue;
        const chain = Array.from(familyIds);
        for (const id of scriptIds) if (!chain.includes(id)) chain.push(id);
        for (const id of backstopIds) {
          if (!chain.includes(id)) chain.push(id);
        }
        if (chain.length > 0) fontChains[requirement.key] = chain;
      }
      if (implicit) {
        const ids = [...new Set(Object.values(fontChains).flat())];
        if (ids.length > 0) fontChains[IMPLICIT_FONT_COVERAGE] = ids;
      }
      return {
        fontChains,
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: {
          compatibilityMode: compat?.compatibilityMode ?? 12,
          noLeading: compat?.noLeading ?? false,
          doNotExpandShiftReturn: compat?.doNotExpandShiftReturn ?? false,
        },
        authoritativeShaping: true,
      };
    },

    coverageCheck(requirements) {
      const chains = new Map<string, FontCoverage[] | null>();
      const settled = new Map<string, Set<string>>();
      for (const requirement of mergeRequirements(requirements)) {
        settled.set(requirement.key, registry.settledBaseLatinCharacters(requirement.key));
        const familyIds = registry.getCachedFontIdChain(requirement.family, requirement.bold, requirement.italic, false);
        const scriptIds = implicit && hosted?.has(requirement.family.trim().toLowerCase()) ? [] :
          registry.getCachedScriptFallbackIds(requirement.scripts ?? [], false);
        if (familyIds === undefined || scriptIds === undefined) continue;
        const earlier = [...familyIds, ...scriptIds];
        const coverage = earlier.map(id => registry.coverageOf(id));
        if (hostedCoverageUnknown(requirement.family) || earlier.length === 0 || coverage.some(covers => !covers)) {
          chains.set(requirement.key, null);
          continue;
        }
        const backstopIds = registry.getCachedBaseLatinIds(requirement.family, requirement.bold, requirement.italic,
          requirement.characters, earlier, requirement.key);
        if (backstopIds === undefined) continue;
        chains.set(requirement.key, [...coverage, ...backstopIds.map(id => registry.coverageOf(id))]
          .filter((covers): covers is FontCoverage => covers !== undefined));
      }
      return next => mergeRequirements(next).every(requirement => {
        const coverage = chains.get(requirement.key);
        if (coverage === undefined) return false;
        const aliases = symbolFamily(requirement.family);
        const known = settled.get(requirement.key);
        return coverage === null || (requirement.characters ?? []).every(character => known?.has(character) ||
          !baseLatinCandidate(character, false) || (baseLatinAlias(character) && !aliases) ||
          coverage.some(covers => covers(character)));
      });
    },

    clear(): void {
      generation++;
      implicit = false;
      registry.clear();
    },
  };
}
