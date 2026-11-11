// @betteroffice/docx declaration yrs/dirtyProjectionStories.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocxProposalResult } from './proposals';
import type { YrsToDocumentOptions } from './yrsToDocument';
/** @internal */
export declare function dirtyProjectionStory(activeStory: string): string;
/** @internal */
export declare class DirtyProjectionStories {
    private readonly stories;
    private nextMark;
    add(story: string): void;
    projectionOptions(): YrsToDocumentOptions | undefined;
    /** The marked stories by their own ids, nested ones included, as the projection memo keys them. */
    capture(): {
        stories: string[];
        clear(): void;
    };
    clear(): void;
}
/**
 * @internal
 * Dirty stories of an editor that keeps a main-thread projection cache and saves in the worker.
 * A worker save projects the stories the main-thread save would project at the same point.
 */
export declare class EditorDirtyStories {
    /** Stories changed since the main-thread projection cache last advanced. */
    readonly projection: DirtyProjectionStories;
    private readonly workerSave;
    private projections;
    private savedAtProjection;
    private adoptingWorkerSaveUpdates;
    add(story: string): void;
    /** Adopts a worker save's own updates without marking stories. */
    adoptWorkerSaveUpdates(apply: () => void): void;
    /** The main-thread projection cache advanced over `projection`'s stories. */
    projected(): void;
    captureWorkerSave(): {
        stories: string[];
        clear(): void;
    };
    clear(): void;
}
/**
 * @internal
 * Runs one session's worker saves in call order. Each save captures its stories once the
 * previous save settled and clears them on success, as consecutive main-thread saves project.
 */
export declare function serialWorkerSaves(dirty: EditorDirtyStories): <T>(save: (stories: string[]) => Promise<T>) => Promise<T>;
/** @internal */
export declare function proposalProjectionStories(known: ReadonlySet<string>, result: Extract<DocxProposalResult, {
    ok: true;
}>, changedStories: readonly string[]): Set<string>;
