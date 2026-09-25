/**
 * Shapes shared by the formats' read-only structured exports. Content trees, anchors, story or
 * sheet selectors, diagnostic codes and limits stay format-owned; these only fix how every format
 * reports diagnostics, truncation and Markdown markers.
 */

export type ExportSeverity = 'info' | 'warning' | 'error';

/** Something an export omitted or could not represent, anchored where it applies. */
export interface ExportDiagnostic<Code extends string, Anchor> {
  code: Code;
  severity: ExportSeverity;
  anchor: Anchor | null;
  message: string;
}

/** Whether an export stopped at its limits before the end of the content. */
export interface ExportCompletion {
  truncated: boolean;
}

/** A Markdown marker comment and the anchor of the content it precedes. */
export interface MarkdownAnchor<Anchor> {
  marker: string;
  anchor: Anchor;
}
