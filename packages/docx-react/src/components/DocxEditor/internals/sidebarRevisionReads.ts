import {
  createYrsSidebarProjection,
  extractTrackedChangesFromYrs,
  type TrackedChangesResult,
  type YrsSidebarProjection,
  type YrsStorySegmentSource,
} from '@betteroffice/docx/layout/render';
import type { YrsRevisionInfo, YrsSession } from '@betteroffice/docx/yrs';

/**
 * The sidebar's revision reads for one session. Revisions and tracked-change entries depend on
 * the document only, so display-list changes (lazily built pages, scrolling) reuse them until the
 * document version changes.
 */
export class SidebarRevisionReads {
  private reads: {
    session: YrsSession;
    version: string;
    revisions: YrsRevisionInfo[];
    tracked: TrackedChangesResult | null;
  } | null = null;

  private delivered: {
    result: TrackedChangesResult;
    to: (result: TrackedChangesResult) => void;
    session: YrsSession;
    version: string;
  } | null = null;

  /** The session's current version and its revisions. */
  revisions(session: YrsSession): { version: string; revisions: YrsRevisionInfo[] } {
    const version = session.version();
    if (this.reads?.session !== session || this.reads.version !== version) {
      this.reads = { session, version, revisions: session.listRevisions(), tracked: null };
    }
    return { version, revisions: this.reads.revisions };
  }

  /** The tracked-change entries of the revisions last read, and the projection they used. */
  tracked(
    session: YrsSession,
    source?: YrsStorySegmentSource
  ): { tracked: TrackedChangesResult; projection: YrsSidebarProjection } {
    const { revisions } = this.revisions(session);
    const projection = createYrsSidebarProjection(session, source);
    const reads = this.reads!;
    reads.tracked ??= extractTrackedChangesFromYrs(revisions, projection);
    return { tracked: reads.tracked, projection };
  }

  /**
   * Hands `result` to `to` unless this document version already delivered it there. The editor
   * also publishes entries itself after an accept or reject, so an equal result at a later version
   * is delivered again.
   */
  deliver(
    to: ((result: TrackedChangesResult) => void) | undefined,
    result: TrackedChangesResult,
    session: YrsSession,
    version: string
  ): void {
    if (!to) return;
    const delivered = this.delivered;
    if (
      delivered?.result === result &&
      delivered.to === to &&
      delivered.session === session &&
      delivered.version === version
    ) {
      return;
    }
    this.delivered = { result, to, session, version };
    to(result);
  }
}
