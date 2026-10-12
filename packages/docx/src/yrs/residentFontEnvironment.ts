import type { YrsResidentFontRegistration } from './index';

export type ResidentFontFace =
  | { family: string; data: Uint8Array; weight?: string | number; style?: 'normal' | 'italic' | 'oblique' }
  | { family: string; substituteOf: number };

export interface ResidentFontEnvironment {
  faces: ResidentFontFace[];
  chains: Record<string, number[]>;
  families: string[];
}

export interface ResidentFontEnvironmentSnapshot {
  revision: number;
  families: string[];
}

export interface ResidentFontEnvironmentEvent extends ResidentFontEnvironmentSnapshot {
  phase: 'ready' | 'presentation';
  frameEpoch?: number;
}

interface PreparedFonts {
  fonts: YrsResidentFontRegistration[];
  css: FontFace[];
  chains: Record<string, number[]>;
  families: string[];
}

export function validFontEnvironment(value: unknown): value is ResidentFontEnvironment {
  if (!value || typeof value !== 'object') return false;
  const env = value as ResidentFontEnvironment;
  const name = (family: unknown): family is string => typeof family === 'string' && !!family.trim() &&
    !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(family);
  return Object.keys(env).every(key => ['faces', 'chains', 'families'].includes(key)) &&
    Array.isArray(env.faces) && env.faces.length <= 256 && env.faces.every((face, index) => {
      if (!face || typeof face !== 'object' || !name(face.family)) return false;
      if ('substituteOf' in face) return Object.keys(face).every(key => ['family', 'substituteOf'].includes(key)) &&
        Number.isSafeInteger(face.substituteOf) && face.substituteOf >= 0 && face.substituteOf < index;
      return Object.keys(face).every(key => ['family', 'data', 'weight', 'style'].includes(key)) &&
        face.data instanceof Uint8Array && face.data.byteLength >= 12 && face.data.byteLength <= 64 * 1024 * 1024 &&
        (face.weight === undefined || /^(normal|bold|[1-9]\d{0,2}|1000)$/.test(String(face.weight))) &&
        (face.style === undefined || ['normal', 'italic', 'oblique'].includes(face.style));
    }) && !!env.chains && typeof env.chains === 'object' && !Array.isArray(env.chains) &&
    Object.entries(env.chains).every(([key, chain]) => name(key) && Array.isArray(chain) &&
      chain.every(slot => Number.isSafeInteger(slot) && slot >= 0 && slot < env.faces.length)) &&
    Array.isArray(env.families) && env.families.every(name);
}

export class ResidentSessionFonts {
  private css: FontFace[] = [];
  private disposed = false;
  private revision = 0;
  snapshot: ResidentFontEnvironmentSnapshot | undefined;

  constructor(
    private readonly engine: { decodeFont(bytes: Uint8Array): Uint8Array },
    private readonly fontSet?: Pick<FontFaceSet, 'add' | 'delete'>,
    private readonly createFace: ((family: string, bytes: Uint8Array, descriptors: FontFaceDescriptors) => FontFace) | undefined =
      typeof FontFace === 'function' ? (family, bytes, descriptors) => new FontFace(family, bytes.slice().buffer, descriptors) : undefined,
  ) {}

  async prepare(environment: ResidentFontEnvironment): Promise<PreparedFonts> {
    if (this.disposed) throw new Error('Resident font session is disposed');
    if (!validFontEnvironment(environment)) throw new Error('Invalid resident font environment');
    const fonts: YrsResidentFontRegistration[] = [];
    const css: FontFace[] = [];
    for (const face of environment.faces) {
      if ('substituteOf' in face) fonts.push({ substituteOf: face.substituteOf, family: face.family });
      else {
        const bytes = this.engine.decodeFont(face.data);
        fonts.push(bytes);
        if (this.createFace && this.fontSet) css.push(this.createFace(face.family, bytes, {
          weight: String(face.weight ?? 'normal'), style: face.style ?? 'normal',
        }));
      }
    }
    await Promise.all(css.map(face => face.load()));
    if (this.disposed) throw new Error('Resident font session is disposed');
    return { fonts, css, chains: structuredClone(environment.chains), families: [...environment.families] };
  }

  commit(prepared: PreparedFonts, register: (fonts: YrsResidentFontRegistration[], chains: Record<string, number[]>) => void): ResidentFontEnvironmentSnapshot {
    if (this.disposed) throw new Error('Resident font session is disposed');
    const added: FontFace[] = [];
    try {
      for (const face of prepared.css) { this.fontSet?.add(face); added.push(face); }
      register(prepared.fonts, prepared.chains);
    } catch (error) {
      for (const face of added) this.fontSet?.delete(face);
      throw error;
    }
    for (const face of this.css) this.fontSet?.delete(face);
    this.css = prepared.css;
    this.snapshot = { revision: ++this.revision, families: prepared.families };
    return this.snapshot;
  }

  dispose(): void {
    this.disposed = true;
    for (const face of this.css) this.fontSet?.delete(face);
    this.css = [];
  }
}
