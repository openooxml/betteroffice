// @betteroffice/docx declaration yrs/selectionChanges.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsSession } from './index';
export type YrsSelectionCause = 'remote' | 'seed';
export declare function yrsSelectionCause(session: YrsSession): YrsSelectionCause | undefined;
export declare function withYrsSelectionCause<T>(session: YrsSession, cause: YrsSelectionCause | undefined, run: () => T): T;
export declare function recordYrsSelectionCause(session: YrsSession, cause?: YrsSelectionCause | undefined, selection?: import("./index").YrsSelection | null): void;
export declare function subscribeYrsSelection(session: YrsSession, listener: () => void): () => void;
