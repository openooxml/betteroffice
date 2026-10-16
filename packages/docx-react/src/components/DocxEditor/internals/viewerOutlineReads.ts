import type { HeadingInfo } from '@betteroffice/docx/utils';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { readAt } from './viewerReads';

type CollectedHeadings = {
  version: string;
  headings: HeadingInfo[];
  targets: Map<number, { story: string; paraId: string }>;
};

export class ViewerOutlineReads {
  private requested: string | null = null;
  private navigation = 0;
  private collected: CollectedHeadings | null = null;
  private pending: { version: string; value: Promise<CollectedHeadings | null> } | null = null;

  constructor(private readonly read: ResidentEngineWorkerClient['documentRead']) {}

  /** Refreshes at one version share one worker read; a later version supersedes it. */
  async collect(version: string, currentVersion: () => string | null): Promise<HeadingInfo[] | null> {
    this.requested = version;
    if (this.collected?.version === version) return this.collected.headings;
    if (this.pending?.version !== version) {
      this.pending = {
        version,
        value: readAt(this.read, { kind: 'headings', expectVersion: version }).then((reply) => {
          if (reply.status !== 'ok' || !reply.value) return null;
          const targets = new Map<number, { story: string; paraId: string }>();
          const headings = reply.value.map(({ story, paraId, text, level, position }) => {
            const pmPos = Math.max(0, position - 1);
            targets.set(pmPos, { story, paraId });
            return { text, level, pmPos };
          });
          return { version, headings, targets };
        }),
      };
    }
    const pending = this.pending;
    const result = await pending.value;
    if (this.pending === pending) this.pending = null;
    if (!result || this.requested !== version || currentVersion() !== version) return null;
    if (this.collected?.version !== version) this.collected = result;
    return this.collected.headings;
  }

  async navigate(
    pmPos: number,
    currentVersion: () => string | null,
    scroll: (position: number) => void
  ): Promise<void> {
    const navigation = ++this.navigation;
    const heading = this.collected?.targets.get(pmPos);
    const version = currentVersion();
    if (!heading || (version !== null && this.collected?.version === version)) {
      scroll(pmPos);
      return;
    }
    if (version === null) return;
    try {
      const reply = await this.read({ kind: 'navigationTarget', ...heading });
      if (navigation !== this.navigation || reply.version !== version || currentVersion() !== version ||
        typeof reply.value !== 'object' || reply.value === null) return;
      scroll(Math.max(0, reply.value.position - 1));
    } catch {
      return;
    }
  }
}
