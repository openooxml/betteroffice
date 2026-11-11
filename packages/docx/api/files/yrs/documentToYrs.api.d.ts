// @betteroffice/docx declaration yrs/documentToYrs.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Direct Document-to-yrs load projection. */
import type { Document } from '../types/document';
import type { YrsOpeningOptions, YrsSession } from './index';
/**
 * Aggregate cap, in xml bytes, for opaque drawing payloads copied into yrs
 * state while seeding. Opaque markup replays verbatim through raw ops and
 * yrs map values, so uncapped documents multiply memory once per layer.
 * Both seeders enforce this constant and refuse over-budget documents
 * instead of silently dropping content.
 */
export declare const OPAQUE_SEED_BUDGET_BYTES: number;
export declare class OpaqueSeedBudgetError extends Error {
    readonly total: number;
    readonly budget: number;
    constructor(total: number);
}
export declare function normalizeSeedError(error: unknown): unknown;
export declare function assertOpaqueSeedBudget(document: Document): void;
/**
 * Seeds every yrs-owned editable story directly from a parsed Document.
 *
 * Stories are `body`, `hf:{rId}`, `fn:{id}`, `en:{id}`, and recursively generated table
 * cell / block-SDT stories. The target session must not already contain any of
 * those story ids. Seeding starts a new opening; see {@link YrsSession.beginOpening}.
 *
 * @public
 */
export declare function documentToYrs(session: YrsSession, document: Document, options?: YrsOpeningOptions): void;
