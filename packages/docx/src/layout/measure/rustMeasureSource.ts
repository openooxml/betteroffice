import type { CompatibilityFlags } from '../../docx/settingsParser';
import {
  TextMeasureFontRegistry,
  type BundledFontProvider,
  type EmbeddedFaceInput,
  type FontScript,
} from './fontRegistry';

export interface RustTextEngine {
  registerFont(bytes: Uint8Array): number;
  clearFonts(): void;
}

export interface ResidentFontRequirement {
  key: string;
  family: string;
  bold: boolean;
  italic: boolean;
  scripts?: FontScript[];
}

export interface ResidentMeasurementConfig {
  fontChains: Record<string, number[]>;
  defaults: { fontSize: 11; fontFamily: 'Calibri' };
  compat: { noLeading: boolean; doNotExpandShiftReturn: boolean };
  authoritativeShaping: true;
}

export interface RustMeasureSource {
  setEmbeddedFaces(faces: EmbeddedFaceInput[]): void;
  setCompat(flags: CompatibilityFlags | undefined): void;
  prepareFontRequirements(requirements: ResidentFontRequirement[]): Promise<boolean>;
  measurementConfigForRequirements(
    requirements: ResidentFontRequirement[]
  ): ResidentMeasurementConfig | undefined;
  clear(): void;
}

let enginePromise: Promise<RustTextEngine> | null = null;

export function getRustTextEngine(): Promise<RustTextEngine> {
  enginePromise ??= import('../wasm/index').then(async (module) => {
    await module.preloadLayoutWasm();
    return {
      registerFont: module.registerMeasureFont,
      clearFonts: module.clearMeasureFonts,
    };
  });
  return enginePromise;
}

export function createRustMeasureSource(options: {
  engine: RustTextEngine;
  bundled?: BundledFontProvider;
}): RustMeasureSource {
  const registry = new TextMeasureFontRegistry(
    { registerFont: (bytes) => options.engine.registerFont(bytes) },
    { bundled: options.bundled }
  );
  let compat: CompatibilityFlags | undefined;

  return {
    setEmbeddedFaces(faces): void {
      registry.setEmbeddedFaces(faces);
    },

    setCompat(flags): void {
      compat = flags;
    },

    async prepareFontRequirements(requirements): Promise<boolean> {
      const settled = await Promise.all(
        requirements.flatMap((requirement) => [
          (async () => {
            try {
              const cached = registry.getCachedFontIdChain(
                requirement.family,
                requirement.bold,
                requirement.italic
              );
              if (cached !== undefined) return false;
              await registry.getFontIdChain(
                requirement.family,
                requirement.bold,
                requirement.italic
              );
              return true;
            } catch {
              return false;
            }
          })(),
          ...Array.from(new Set(requirement.scripts ?? []), async (script) => {
            try {
              if (registry.getCachedScriptFallbackIds([script]) !== undefined) return false;
              await registry.getScriptFallbackIds([script]);
              return true;
            } catch {
              return false;
            }
          }),
        ])
      );
      return settled.some(Boolean);
    },

    measurementConfigForRequirements(requirements): ResidentMeasurementConfig | undefined {
      const fontChains: Record<string, number[]> = {};
      for (const requirement of requirements) {
        const familyIds = registry.getCachedFontIdChain(
          requirement.family,
          requirement.bold,
          requirement.italic
        );
        const scriptIds = registry.getCachedScriptFallbackIds(requirement.scripts ?? []);
        if (familyIds === undefined || scriptIds === undefined) return undefined;
        const chain = Array.from(familyIds);
        for (const id of scriptIds) if (!chain.includes(id)) chain.push(id);
        if (chain.length > 0) fontChains[requirement.key] = chain;
      }
      return {
        fontChains,
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: {
          noLeading: compat?.noLeading ?? false,
          doNotExpandShiftReturn: compat?.doNotExpandShiftReturn ?? false,
        },
        authoritativeShaping: true,
      };
    },

    clear(): void {
      registry.clear();
    },
  };
}
