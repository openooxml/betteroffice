import type { YrsResidentFontRegistration } from './index';
import type { ProgressiveFontSpec } from './residentProgressiveFonts';

export type DocxFontFace =
  | { family: string; data: Uint8Array; weight?: string | number; style?: 'normal' | 'italic' | 'oblique' }
  | { family: string; substituteOf: number };

export interface DocxFontEnvironment {
  faces: DocxFontFace[];
  chains: Record<string, number[]>;
  families: string[];
  /** Faces registered from a sliced face's skeleton and core, whose outlines the worker loads as pages need them. */
  progressive?: ProgressiveFontSpec[];
}

export interface DocxFontEnvironmentSnapshot {
  revision: number;
  families: string[];
}

export interface DocxFontEnvironmentEvent extends DocxFontEnvironmentSnapshot {
  phase: 'ready' | 'presentation';
  frameEpoch?: number;
}

interface PreparedFonts {
  fonts: YrsResidentFontRegistration[];
  css: FontFace[];
  chains: Record<string, number[]>;
  families: string[];
  progressive?: ProgressiveFontSpec[];
}

const text = (value: unknown, max = 4096): value is string => typeof value === 'string' && value.length > 0 && value.length <= max;
const hex = (value: unknown): value is string => typeof value === 'string' && /^[0-9a-f]{64}$/.test(value);

function validProgressiveSpecs(specs: unknown, faces: readonly DocxFontFace[]): specs is ProgressiveFontSpec[] {
  if (!Array.isArray(specs) || specs.length > faces.length) return false;
  const slots = new Set<number>();
  return specs.every((spec: ProgressiveFontSpec) => {
    if (!spec || typeof spec !== 'object' || !Number.isSafeInteger(spec.slot) || spec.slot < 0 || spec.slot >= faces.length ||
      slots.has(spec.slot) || 'substituteOf' in faces[spec.slot]!) return false;
    slots.add(spec.slot);
    const source = spec.source;
    if (!source || typeof source !== 'object' || !source.ref || typeof source.ref !== 'object') return false;
    const { ref, urls, timeoutMs } = source;
    return Number.isSafeInteger(ref.format) && text(ref.index) && hex(ref.indexSha256) && hex(ref.sha256) &&
      Number.isSafeInteger(ref.byteLength) && ref.byteLength > 0 &&
      !!urls && typeof urls === 'object' && !Array.isArray(urls) && Object.keys(urls).length <= 4096 &&
      Object.entries(urls).every(([file, url]) => text(file) && text(url, 8192)) &&
      (timeoutMs === undefined || (Number.isSafeInteger(timeoutMs) && timeoutMs > 0));
  });
}

export function validFontEnvironment(value: unknown): value is DocxFontEnvironment {
  if (!value || typeof value !== 'object') return false;
  const env = value as DocxFontEnvironment;
  const name = (family: unknown): family is string => typeof family === 'string' && !!family.trim() &&
    !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(family);
  return Object.keys(env).every(key => ['faces', 'chains', 'families', 'progressive'].includes(key)) &&
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
    Array.isArray(env.families) && env.families.every(name) &&
    (env.progressive === undefined || validProgressiveSpecs(env.progressive, env.faces));
}

/** The characters docx-layout's `paragraph_scripts` gives a script fallback face. */
const FALLBACK_SCRIPT_TEXT = /[\u0590-\u06ff\u0750-\u077f\u0870-\u08ff\u1100-\u11ff\u3000-\u30ff\u3130-\u318f\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff\ua960-\ua97f\uac00-\ud7ff\uf900-\ufaff\ufb1d-\ufdff\ufe30-\ufe4f\ufe70-\ufeff\uff00-\uffdc\u{20000}-\u{3ffff}]/u;
/**
 * Commands that can bring back or reveal text, or change the fonts runs use, without naming
 * either. A remote merge can also integrate updates held back for a missing dependency.
 */
const FONT_RESHAPING_COMMANDS = new Set(['pasteFragment', 'undo', 'redo', 'toggleMark', 'formatSelection', 'formatRange',
  'clearFormatting', 'applyParagraphStyle', 'acceptRevision', 'rejectRevision', 'setProposalStates', 'applyRemoteUpdate',
  'applyRemoteUpdateChunks', 'collabPipe']);

/** Fallback-script text in a string field, or a paragraph style set by a step such as `setParagraphStyle`. */
function carriesFallbackScript(value: unknown, depth = 0): boolean {
  if (typeof value === 'string') return FALLBACK_SCRIPT_TEXT.test(value);
  if (value instanceof ArrayBuffer || ArrayBuffer.isView(value)) return false;
  if (!value || typeof value !== 'object' || depth > 8) return false;
  if (!Array.isArray(value) && typeof (value as { styleId?: unknown }).styleId === 'string') return true;
  return (Array.isArray(value) ? value : Object.values(value)).some(item => carriesFallbackScript(item, depth + 1));
}

/**
 * Whether an applied command may need fonts the open-time requirements lacked: text in a script
 * with its own fallback face, a paragraph style it sets, or a change the payload does not spell out.
 * Showing or hiding hidden text is judged by the worker, which knows the setting before and after.
 */
export function commandMayNeedNewFonts(kind: string, payload: unknown): boolean {
  if (kind === 'setFontEnvironment' || kind === 'setRenderEnvironment') return false;
  return FONT_RESHAPING_COMMANDS.has(kind) || carriesFallbackScript(payload);
}

export class ResidentSessionFonts {
  private css: FontFace[] = [];
  private disposed = false;
  private revision = 0;
  snapshot: DocxFontEnvironmentSnapshot | undefined;

  constructor(
    private readonly engine: { decodeFont(bytes: Uint8Array): Uint8Array },
    private readonly fontSet?: Pick<FontFaceSet, 'add' | 'delete'>,
    private readonly createFace: ((family: string, bytes: Uint8Array, descriptors: FontFaceDescriptors) => FontFace) | undefined =
      typeof FontFace === 'function' ? (family, bytes, descriptors) => new FontFace(family, bytes.slice().buffer, descriptors) : undefined,
  ) {}

  async prepare(environment: DocxFontEnvironment): Promise<PreparedFonts> {
    if (this.disposed) throw new Error('Worker font session is disposed');
    if (!validFontEnvironment(environment)) throw new Error('Invalid worker font environment');
    const fonts: YrsResidentFontRegistration[] = [];
    const css: FontFace[] = [];
    const sliced = new Set(environment.progressive?.map(spec => spec.slot));
    for (const [slot, face] of environment.faces.entries()) {
      if ('substituteOf' in face) fonts.push({ substituteOf: face.substituteOf, family: face.family });
      else {
        const bytes = this.engine.decodeFont(face.data);
        fonts.push(bytes);
        if (this.createFace && this.fontSet && !sliced.has(slot)) css.push(this.createFace(face.family, bytes, {
          weight: String(face.weight ?? 'normal'), style: face.style ?? 'normal',
        }));
      }
    }
    await Promise.all(css.map(face => face.load()));
    if (this.disposed) throw new Error('Worker font session is disposed');
    return { fonts, css, chains: structuredClone(environment.chains), families: [...environment.families],
      ...(environment.progressive?.length ? { progressive: structuredClone(environment.progressive) } : {}) };
  }

  commit(prepared: PreparedFonts, register: (fonts: YrsResidentFontRegistration[], chains: Record<string, number[]>,
    progressive: ProgressiveFontSpec[] | undefined) => void): DocxFontEnvironmentSnapshot {
    if (this.disposed) throw new Error('Worker font session is disposed');
    const added: FontFace[] = [];
    try {
      for (const face of prepared.css) { this.fontSet?.add(face); added.push(face); }
      register(prepared.fonts, prepared.chains, prepared.progressive);
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
