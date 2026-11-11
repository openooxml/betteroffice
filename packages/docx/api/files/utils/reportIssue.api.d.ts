// @betteroffice/docx declaration utils/reportIssue.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Framework-agnostic helper to open a pre-filled GitHub "new issue"
 * URL for the Help > Report issue action. GitHub doesn't allow file
 * attachments via URL params, so the body asks the user to drag-drop
 * their DOCX onto the issue after it opens.
 *
 * Lifted from packages/docx-react/src/components/reportIssue.ts so the
 * Vue adapter can re-use it without re-implementing.
 * @packageDocumentation
 * @public
 */
export interface ReportIssueEnv {
    userAgent?: string;
    viewport?: {
        width: number;
        height: number;
    };
    pageUrl?: string;
}
export declare function buildReportIssueUrl(env?: ReportIssueEnv): string;
export declare function openReportIssue(env?: ReportIssueEnv): void;
