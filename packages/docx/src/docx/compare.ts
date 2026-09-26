/**
 * Compares two DOCX packages into tracked changes. Only text in body paragraphs of unchanged
 * structure is compared; anything else refuses with diagnostics, never a partial redline.
 */

/** Bounds a comparison enforces. The defaults are the v1 ceilings; callers may only tighten them. */
export interface DocxCompareLimits {
  /** Bytes of each input package. Default 32 MiB. */
  maxInputBytes: number;
  /** Inflated bytes of each input package. Default 128 MiB. */
  maxExpandedBytes: number;
  /** Paragraphs of each input, every story counted. Default 10,000. */
  maxParagraphs: number;
  /** UTF-16 units of text in every story of both inputs combined. Default 1,048,576. */
  maxTextUnits: number;
  /** Paragraph-alignment table cells, cumulative. Default 250,000. */
  maxAlignmentCells: number;
  /** Text-diff table cells, cumulative, similarity checks included. Default 4,000,000. */
  maxDiffCells: number;
  /** Changes, which is also the number of batch steps. Default 128. */
  maxChanges: number;
  /** Diagnostics; one more ends inspection with `diagnostics-truncated`. Default 256. */
  maxDiagnostics: number;
  /** Bytes of the encoded editing state staged for the batch. Default 64 MiB. */
  maxStagedBytes: number;
  /** Bytes of the result as JSON, without the output package; at least 1,024. Default 8 MiB. */
  maxResultBytes: number;
  /** Bytes of the output package, the unchanged original included. Default 64 MiB. */
  maxOutputBytes: number;
}

export interface DocxCompareOptions {
  /** The author of every revision. Must not be blank or contain control characters. */
  author: string;
  /** RFC 3339 timestamp with an explicit offset, recorded in UTC. The clock is never read. */
  date: string;
  /** Unicode words, keeping punctuation and whitespace tokens (default), or grapheme clusters. */
  granularity?: 'word' | 'char';
  /**
   * `fail` (default) stops at the first blocking condition; `report` keeps inspecting and
   * returns every diagnostic. Neither returns a document while a blocking condition exists.
   */
  unsupported?: 'fail' | 'report';
  limits?: Partial<DocxCompareLimits>;
}

/**
 * Half-open UTF-16 range in the projected text of one input's body paragraph (tabs are `\t`,
 * other inline atoms one U+FFFC). `path` is the paragraph's element-child path from the root of
 * `part`, whose SHA-256 is `partSha256`. Spans address the inputs, not the output.
 */
export interface DocxCompareTextSpan {
  part: string;
  partSha256: string;
  path: number[];
  start: number;
  end: number;
  text: string;
}

/**
 * One tracked replacement, ordered by original then revised position. An insertion has an empty
 * `original` span and a deletion an empty `revised` span. `id` (`change-0`, …) is local to this
 * result: neither an OOXML revision id nor a durable anchor.
 */
export interface DocxComparedChange {
  id: string;
  kind: 'insertion' | 'deletion' | 'replacement';
  original: DocxCompareTextSpan;
  revised: DocxCompareTextSpan;
}

export interface DocxCompareLocation {
  input: 'original' | 'revised' | 'output';
  part: string | null;
  path: number[] | null;
  start: number | null;
  end: number | null;
}

export type DocxCompareDiagnosticCode =
  | 'invalid-options'
  | 'invalid-docx'
  | 'existing-revisions'
  | 'paragraph-insertion'
  | 'paragraph-deletion'
  | 'paragraph-move'
  | 'ambiguous-alignment'
  | 'table-change'
  | 'structure-change'
  | 'field-change'
  | 'object-change'
  | 'content-control-change'
  | 'formatting-change'
  | 'unsupported-content'
  | 'unsupported-formatting'
  | 'out-of-scope-change'
  | 'opaque-part-change'
  | 'provenance-unavailable'
  | 'ambiguous-identity'
  | 'metadata-difference'
  | 'limit-exceeded'
  | 'diagnostics-truncated'
  | 'batch-refused'
  | 'serialization-failed'
  | 'roundtrip-mismatch';

/** `error` diagnostics block the comparison; `info` and `warning` ones accompany a result. */
export interface DocxCompareDiagnostic {
  code: DocxCompareDiagnosticCode;
  severity: 'info' | 'warning' | 'error';
  message: string;
  locations: DocxCompareLocation[];
}

/**
 * A redline, or a refusal as data; with no differences `docx` is the original bytes. Malformed
 * options and internal failures throw.
 */
export type DocxCompareResult =
  | {
      ok: true;
      docx: Uint8Array;
      changes: DocxComparedChange[];
      diagnostics: DocxCompareDiagnostic[];
    }
  | {
      ok: false;
      diagnostics: DocxCompareDiagnostic[];
    };

/**
 * Compares two DOCX packages and returns the original with the revised body text differences
 * as tracked changes by `options.author` at `options.date`.
 */
export async function compareDocx(
  original: Uint8Array,
  revised: Uint8Array,
  options: DocxCompareOptions
): Promise<DocxCompareResult> {
  const { compareDocxInSession } = await import('../yrs/compareDocx');
  return compareDocxInSession(original, revised, options);
}
