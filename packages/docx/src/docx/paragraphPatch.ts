export type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
export type Attrs = Record<string, Json>;
export type SlotV1 = { state: 'missing' } | { state: 'present'; value: Json };
export type AttrDeltaV1 = { path: string[]; before: SlotV1; after: SlotV1 };
export type UnitV1 =
  | { kind: 'text'; text: string; attrs: Attrs }
  | { kind: 'atom'; atomKind: string; payload: Attrs; attrs: Attrs };
export type InheritanceV1 = { paragraph: Attrs; runs: Attrs[]; paragraphMark: Attrs };
export type SeedV1 = {
  units: UnitV1[];
  properties: Attrs;
  unitSha256: string;
  inherit?: InheritanceV1;
};
export type DonorV1 = { kind: 'at'; at: number } | { kind: 'pilcrow' };
export type InsertedTextV1 = { text: string; attrs: Attrs; inherit: Attrs };
export type TextEditV1 = { from: number; to: number; insert: InsertedTextV1[]; donor: DonorV1 };
export type RunEditV1 = {
  from: number;
  to: number;
  deltas: AttrDeltaV1[];
  target: Attrs;
  inherit: Attrs;
};
export type ParagraphPatchV1 = {
  ordinal: number;
  sessionKey: string;
  seed: SeedV1;
  textEdits: TextEditV1[];
  runEdits: RunEditV1[];
  paraDeltas: AttrDeltaV1[];
  paraInherit: Attrs;
  markInherit?: Attrs;
  resultSha256: string;
};
export type PatchErrorCode =
  | 'MISSING_SOURCE'
  | 'SOURCE_DIGEST_MISMATCH'
  | 'INVALID_PLAN'
  | 'SEED_MISMATCH'
  | 'AMBIGUOUS_ALIGNMENT'
  | 'AMBIGUOUS_DONOR'
  | 'INVALID_BOUNDARY'
  | 'UNSUPPORTED_KEY'
  | 'UNSUPPORTED_STRUCTURE'
  | 'PROPERTY_CONFLICT'
  | 'PATCH_CONFLICT'
  | 'RESULT_MISMATCH';
export type LegacyParagraphV1 = { ordinal: number; code: PatchErrorCode; detail: string };
