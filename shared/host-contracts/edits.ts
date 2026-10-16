/**
 * Result envelopes shared by the formats' host edit APIs. Targets, steps, receipts and failure
 * codes stay format-owned; these shapes only fix how every format reports success and refusal.
 */

/** Why a host operation was refused, as data. The document is unchanged. */
export interface OperationFailure<Code extends string, Target> {
  code: Code;
  /** The request step that failed, when one did. */
  stepIndex?: number;
  /** The earlier step an `overlapping-steps` failure conflicts with. */
  conflictingStepIndex?: number;
  /** The failing step's target as requested or resolved. */
  target?: Target;
  message: string;
}

/** A policy refusal carrying the version the document is still at. */
export interface OperationRefusal<Failure> {
  ok: false;
  version: string;
  failure: Failure;
}

/**
 * An accepted edit request. `applied` is false when every step was a no-op; the version then
 * stays `baseVersion` and no history or change notification was produced.
 */
export interface EditSuccess<Receipt> {
  ok: true;
  baseVersion: string;
  version: string;
  applied: boolean;
  /** One receipt per request step, in request order. */
  receipts: Receipt[];
}

/** A validated edit request. Validation reserves nothing and cannot be committed later. */
export interface ValidationSuccess<Preview> {
  ok: true;
  baseVersion: string;
  wouldApply: boolean;
  previews: Preview[];
}
