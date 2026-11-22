type Failure = { error: Error; until: number };

export interface ImplicitFontCdnBase {
  transfer: Promise<void>;
  unavailable?: Failure;
  invalid: Map<string, Failure>;
  warnedUntil: number;
}

const COOLDOWN = Symbol.for('betteroffice.fonts.cdn.cooldown');
const scope = globalThis as typeof globalThis & { [COOLDOWN]?: Map<string, ImplicitFontCdnBase> };

export function implicitFontCdnBases(): Map<string, ImplicitFontCdnBase> {
  return scope[COOLDOWN] ??= new Map();
}
