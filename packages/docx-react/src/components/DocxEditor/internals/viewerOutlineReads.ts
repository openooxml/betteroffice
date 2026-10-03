import type { HeadingInfo } from '@betteroffice/docx/utils';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { readAt } from './viewerReads';

export class ViewerOutlineReads {
  private refresh = 0;
  private navigation = 0;
  private collected: { version: string; headings: HeadingInfo[] } | null = null;
  private targets = new Map<number, { story: string; paraId: string }>();

  constructor(private readonly read: ResidentEngineWorkerClient['documentRead']) {}

  async collect(version: string, currentVersion: () => string | null): Promise<HeadingInfo[] | null> {
    const refresh = ++this.refresh;
    if (this.collected?.version === version) return this.collected.headings;
    const reply = await readAt(this.read, { kind: 'headings', expectVersion: version });
    if (refresh !== this.refresh || currentVersion() !== version ||
      reply.status !== 'ok' || !reply.value) return null;
    this.targets.clear();
    const headings = reply.value.map(({ story, paraId, text, level, position }) => {
      const pmPos = Math.max(0, position - 1);
      this.targets.set(pmPos, { story, paraId });
      return { text, level, pmPos };
    });
    this.collected = { version, headings };
    return headings;
  }

  async navigate(
    pmPos: number,
    currentVersion: () => string | null,
    scroll: (position: number) => void
  ): Promise<void> {
    const navigation = ++this.navigation;
    const heading = this.targets.get(pmPos);
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
