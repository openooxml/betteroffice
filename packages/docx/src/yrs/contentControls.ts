/**
 * Content-control discovery: every control of a document in document order, with the metadata the
 * structured export reports, its canonical text, placement, anchor, nesting and effective lock.
 * Plain JSON, produced by one Rust inventory for sessions, bytes, the native facade and Python.
 *
 * `controlId` is an opaque engine locator scoped to the version (or snapshot) it was read at; it
 * does not survive save and reopen. An inline control's id can name a different control after
 * edits, so list again after a version change or select by tag. Fill controls with a
 * `setContentControlText` edit step; concurrent fills from collaborators resolve last-writer-wins
 * for an inline control and merge like concurrent typing for a block control.
 */

import type { OperationRefusal } from '../../../../shared/host-contracts/edits';
import type { ExportDiagnostic } from '../../../../shared/host-contracts/exports';
import type { DocxAnchor, DocxControlMetadata, DocxStorySelection } from './readTypes';
import type { DocxExportDiagnosticCode, DocxExportFailure } from './structuredExport';

/**
 * A control's canonical displayed text: tabs stay tabs, line breaks and paragraph boundaries read
 * as LF. `unavailable` when the content has no faithful text projection or the session cannot
 * read it.
 */
export type DocxContentControlValue =
  | { kind: 'text'; text: string }
  | {
      kind: 'unavailable';
      reason:
        | 'non-text-content'
        | 'tracked-revisions'
        | 'provenance-unavailable'
        | 'unsupported-story';
    };

/**
 * One content control. `lock` is the control's own authored lock; `effectiveLock` combines it
 * with every control containing it, and `known` is false when any of them carries a lock value
 * the engine does not know.
 */
export interface DocxContentControl extends DocxControlMetadata {
  placement: 'inline' | 'block';
  anchor: DocxAnchor;
  parentControlId: string | null;
  value: DocxContentControlValue;
  /** Whether a plain-text control accepts line breaks (`w:multiLine`); `null` for other types. */
  multiLine: boolean | null;
  effectiveLock: {
    content: boolean;
    control: boolean;
    known: boolean;
  };
}

/**
 * What to list. `stories` defaults to every category, `maxControls` to 10,000 (at most
 * 1,000,000) and `maxBytes` to 8,388,608 (1,024 to 67,108,864, measured on the compact JSON).
 * Exceeding a limit refuses with `limit-exceeded` rather than returning a prefix.
 */
export interface DocxContentControlsOptions {
  stories?: readonly DocxStorySelection[];
  maxControls?: number;
  maxBytes?: number;
}

/** The control a write addresses: its engine id, or its tag, which must then be unique. */
export type DocxContentControlSelector =
  | { kind: 'id'; controlId: string }
  | { kind: 'tag'; tag: string };

/** Which controls a find returns: every exact, case-sensitive match, none trimmed. */
export type DocxContentControlQuery = DocxContentControlSelector | { kind: 'alias'; alias: string };

export type DocxContentControlDiagnostic = ExportDiagnostic<DocxExportDiagnosticCode, DocxAnchor>;

/**
 * The controls of a document in document order: body, headers, footers, footnotes, endnotes,
 * comments, each control before its descendants. `complete` is false when diagnostics name a
 * coverage gap. Controls kept only as source XML (raw blocks, comment bodies) are never writable;
 * they carry `sourcePart` anchors while the package holds their current XML, and a raw block
 * changed since opening anchors its controls to the block instead.
 */
export interface DocxContentControlsSnapshot {
  schemaVersion: 1;
  anchorScope: 'session' | 'snapshot';
  includedStories: DocxStorySelection[];
  controls: DocxContentControl[];
  complete: boolean;
  diagnostics: DocxContentControlDiagnostic[];
}

/** Options out of range or above a hard maximum, or a session holding no document content. */
export type DocxContentControlReadFailure = DocxExportFailure;

/** A session read: the controls with the version they were read at, or a refusal. */
export type DocxContentControlsResult =
  | { ok: true; version: string; content: DocxContentControlsSnapshot }
  | OperationRefusal<DocxContentControlReadFailure>;
