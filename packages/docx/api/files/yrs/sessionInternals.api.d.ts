// @betteroffice/docx declaration yrs/sessionInternals.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Engine operations the package runs on sessions it creates itself; not public API. */
import type { editorSaveKeys, noteSeparatorKeys } from './editorSaveKeys';
import type { YrsSession } from './index';
import type { Document } from '../types/document';
export interface YrsSessionInternals {
    headerFooterAliases(): ReadonlyMap<string, string>;
    setHeaderFooterAliases(json: string): void;
    sourcePackage(): {
        buffer: ArrayBuffer;
        keys: ReturnType<typeof editorSaveKeys>;
    } | null;
    materializedSource(): Readonly<Document> | null;
    /** The `data:` sources this session seeded, as `media:{n}` tokens; empty when none. */
    mediaSources(): string;
    /** The opened source's note separators, keyed as `noteSeparatorKeys` does. */
    sourceNoteSeparators(): ReturnType<typeof noteSeparatorKeys> | null;
    /** Compares two DOCX packages into the empty session; bridge JSON. */
    compareDocx(original: Uint8Array, revised: Uint8Array, options: string): string;
    /** The final comparison result for the saved bytes; bridge JSON. */
    finishComparedDocx(bytes: Uint8Array): string;
    /** The final comparison result when saving failed; bridge JSON. */
    failComparedDocx(message: string): string;
}
/** Returns the session's opened source package when available. @internal */
export declare function sessionSourcePackage(session: YrsSession): {
    buffer: ArrayBuffer;
    keys: ReturnType<typeof editorSaveKeys>;
} | null;
export declare function sessionMaterializedSource(session: YrsSession): Readonly<Document> | null;
/** The opened source's note separator keys when available. @internal */
export declare function sessionSourceNoteSeparators(session: YrsSession): ReturnType<typeof noteSeparatorKeys> | null;
export declare function registerSessionInternals(session: object, internals: YrsSessionInternals): void;
export declare function sessionInternals(session: object): YrsSessionInternals;
/** @internal */
export declare function headerFooterStory(session: object | null | undefined, rId: string): string;
/** The relationship ids whose bands show `rId`'s story. @internal */
export declare function headerFooterDisplayRIds(session: object, rId: string): string[];
/** `rId`'s canonical relationship id and every alias `aliases` routes to it. @internal */
export declare function routedHeaderFooterRIds(aliases: ReadonlyMap<string, string> | undefined, rId: string): string[];
/** The `data:` image sources the session seeded, as its media tokens. @internal */
export declare function sessionMediaSources(session: YrsSession): string;
