// @betteroffice/docx declaration yrs/residentFontEnvironment.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsResidentFontRegistration } from './index';
import type { ProgressiveFontSpec } from './residentProgressiveFonts';
export type DocxFontFace = {
    family: string;
    data: Uint8Array;
    weight?: string | number;
    style?: 'normal' | 'italic' | 'oblique';
} | {
    family: string;
    substituteOf: number;
};
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
export declare function validFontEnvironment(value: unknown): value is DocxFontEnvironment;
/**
 * Whether an applied command may need fonts the open-time requirements lacked: text in a script
 * with its own fallback face, a paragraph style it sets, or a change the payload does not spell out.
 * Showing or hiding hidden text is judged by the worker, which knows the setting before and after.
 */
export declare function commandMayNeedNewFonts(kind: string, payload: unknown): boolean;
export declare class ResidentSessionFonts {
    private readonly engine;
    private readonly fontSet?;
    private readonly createFace;
    private css;
    private disposed;
    private revision;
    snapshot: DocxFontEnvironmentSnapshot | undefined;
    constructor(engine: {
        decodeFont(bytes: Uint8Array): Uint8Array;
    }, fontSet?: Pick<FontFaceSet, "add" | "delete"> | undefined, createFace?: ((family: string, bytes: Uint8Array, descriptors: FontFaceDescriptors) => FontFace) | undefined);
    prepare(environment: DocxFontEnvironment): Promise<PreparedFonts>;
    commit(prepared: PreparedFonts, register: (fonts: YrsResidentFontRegistration[], chains: Record<string, number[]>, progressive: ProgressiveFontSpec[] | undefined) => void): DocxFontEnvironmentSnapshot;
    dispose(): void;
}
export {};
