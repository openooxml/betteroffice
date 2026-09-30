import type { YrsStorySegmentSource } from '@betteroffice/docx/layout/render';
import type { YrsSession, YrsStorySegment } from '@betteroffice/docx/yrs';

interface CachedStory {
  /** Per-paragraph digests; null until known. */
  digests: string[] | null;
  segments: YrsStorySegment[];
}

/**
 * `storySegments` for one session, re-read only where the document changed:
 * {@link YrsStorySegmentCache.refresh} drops the stories changed since the
 * last refresh. A dropped story whose paragraph digests are known re-reads
 * only the paragraphs whose segments differ from ones already held; digests
 * of stories read whole are fetched while the main thread is idle. Segments
 * are shared across reads, so callers must not mutate them.
 */
export class YrsStorySegmentCache {
  private revision = 0;
  private readonly stories = new Map<string, CachedStory>();
  private readonly stale = new Map<string, CachedStory>();
  private readonly units = new Map<string, { segments: YrsStorySegment[]; stories: number }>();
  private readonly undigested = new Set<string>();
  private cancelIdle: (() => void) | null = null;
  private released = false;

  constructor(readonly session: YrsSession) {}

  /**
   * Brings the cache up to the session's current state. A changed story's
   * previous paragraphs stay available to its next read until the refresh
   * after that.
   */
  refresh(): void {
    const { revision, stories } = this.session.storiesChangedSince(this.revision);
    if (revision === this.revision && stories.length === 0) return;
    for (const replaced of this.stale.values()) this.release(replaced.digests ?? []);
    this.stale.clear();
    this.revision = revision;
    for (const story of stories) {
      const cached = this.stories.get(story);
      if (!cached) continue;
      this.stories.delete(story);
      this.stale.set(story, cached);
    }
  }

  /**
   * `storySegments(story)`, reading only paragraphs this cache does not hold.
   * Current for the stories changed before the last refresh.
   */
  segments(story: string): YrsStorySegment[] {
    const cached = this.stories.get(story);
    if (cached) return cached.segments;
    if (!this.stale.get(story)?.digests) {
      const segments = this.session.storySegments(story);
      this.store(story, null, [segments]);
      if (splitUnits(segments).length > 1) this.undigested.add(story);
      return segments;
    }
    const digests = this.session.storySegmentUnitDigests(story);
    const missing = digests.flatMap((digest, index) => (this.units.has(digest) ? [] : [index]));
    const fetched = missing.length > 0 ? this.session.storySegmentUnits(story, missing) : [];
    const units = digests.map((digest) => this.units.get(digest)?.segments ?? []);
    missing.forEach((index, fetchedIndex) => {
      units[index] = fetched[fetchedIndex]!;
    });
    return this.store(story, digests, units);
  }

  /** Fetches the digests of stories read whole once the main thread is idle. */
  scheduleDigests(): void {
    if (this.cancelIdle || this.undigested.size === 0) return;
    const run = () => {
      this.cancelIdle = null;
      try {
        this.completeDigests();
      } catch {
        // A session destroyed meanwhile has nothing left to digest.
      }
    };
    if (typeof requestIdleCallback === 'function') {
      const id = requestIdleCallback(run);
      this.cancelIdle = () => cancelIdleCallback(id);
    } else {
      const id = setTimeout(run, 0);
      this.cancelIdle = () => clearTimeout(id);
    }
  }

  /** Whether {@link YrsStorySegmentCache.dispose} ran; a disposed cache holds nothing. */
  get disposed(): boolean {
    return this.released;
  }

  dispose(): void {
    this.cancelIdle?.();
    this.cancelIdle = null;
    this.released = true;
    this.stories.clear();
    this.stale.clear();
    this.units.clear();
    this.undigested.clear();
  }

  /** Reads the digests of stories read whole that have not changed since. */
  completeDigests(): void {
    const { stories: changed } = this.session.storiesChangedSince(this.revision);
    const changedSince = new Set(changed);
    for (const story of this.undigested) {
      const cached = this.stories.get(story);
      if (!cached || cached.digests || changedSince.has(story)) continue;
      const digests = this.session.storySegmentUnitDigests(story);
      const units = splitUnits(cached.segments);
      if (digests.length !== units.length) continue;
      this.stories.delete(story);
      this.store(story, digests, units);
    }
    this.undigested.clear();
  }

  private store(
    story: string,
    digests: string[] | null,
    units: YrsStorySegment[][]
  ): YrsStorySegment[] {
    digests?.forEach((digest, index) => {
      const unit = this.units.get(digest);
      if (unit) unit.stories += 1;
      else this.units.set(digest, { segments: units[index]!, stories: 1 });
    });
    const replaced = this.stale.get(story);
    if (replaced) {
      this.stale.delete(story);
      this.release(replaced.digests ?? []);
    }
    const segments = digests
      ? digests.flatMap((digest) => this.units.get(digest)!.segments)
      : units.flat();
    this.stories.set(story, { digests, segments });
    return segments;
  }

  private release(digests: readonly string[]): void {
    for (const digest of digests) {
      const unit = this.units.get(digest);
      if (unit && --unit.stories === 0) this.units.delete(digest);
    }
  }
}

/**
 * Serves `session`'s segments from `cache`, brought up to date on every read. Once the cache is
 * disposed, as when its editor unmounts or opens another document, reads go to `session` itself.
 */
export function storySegmentSource(
  session: YrsSession,
  cache: YrsStorySegmentCache
): YrsStorySegmentSource {
  return {
    segments(story) {
      if (cache.disposed || cache.session !== session) return session.storySegments(story);
      cache.refresh();
      const segments = cache.segments(story);
      cache.scheduleDigests();
      return segments;
    },
  };
}

/** Segments split after each pilcrow, as the session's paragraph digests split them. */
function splitUnits(segments: readonly YrsStorySegment[]): YrsStorySegment[][] {
  const units: YrsStorySegment[][] = [];
  let unit: YrsStorySegment[] = [];
  for (const segment of segments) {
    unit.push(segment);
    if (segment.kind === 'pilcrow') {
      units.push(unit);
      unit = [];
    }
  }
  if (unit.length > 0) units.push(unit);
  return units;
}
