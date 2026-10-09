import type { YrsStorySegmentSource } from '@betteroffice/docx/layout/render';
import type { YrsSession, YrsStorySegment } from '@betteroffice/docx/yrs';
import { scheduleIdleWork } from '../hooks/pageBuildScheduler';

const IDLE_SLICE_MS = 8;
const DIGEST_FALLBACK_MS = 5000;

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
  private readonly warming = new Map<string, CachedStory | null>();
  private cancelIdle: (() => void) | null = null;
  private digestExpiresAt: number | null = null;
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
      if (splitUnits(segments).length > 1) {
        this.undigested.add(story);
        if (this.warming.has(story)) this.warming.set(story, this.stories.get(story)!);
      }
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
    if (this.released || this.cancelIdle || this.undigested.size === 0) return;
    this.digestExpiresAt ??=
      performance.now() + (typeof requestIdleCallback === 'function' ? DIGEST_FALLBACK_MS : 0);
    const run = (deadline?: IdleDeadline) => {
      this.cancelIdle = null;
      try {
        this.completeDigests(deadline ?? {
          didTimeout: true,
          timeRemaining: () => 0,
        });
        if (this.undigested.size > 0) this.scheduleDigests();
        else this.digestExpiresAt = null;
      } catch {
        // A session destroyed meanwhile has nothing left to digest.
      }
    };
    this.cancelIdle = scheduleIdleWork(run, this.digestExpiresAt, 0).cancel;
  }

  /** Whether {@link YrsStorySegmentCache.dispose} ran; a disposed cache holds nothing. */
  get disposed(): boolean {
    return this.released;
  }

  dispose(): void {
    this.cancelIdle?.();
    this.cancelIdle = null;
    this.digestExpiresAt = null;
    this.released = true;
    this.stories.clear();
    this.stale.clear();
    this.units.clear();
    this.undigested.clear();
    this.warming.clear();
  }

  /** Warms pending stories, re-reading those edited after a yield. */
  completeDigests(deadline?: IdleDeadline): void {
    if (this.released) return;
    const start = performance.now();
    const { stories: changed } = this.session.storiesChangedSince(this.revision);
    const changedSince = new Set(changed);
    if (this.warming.size === 0) {
      for (const story of this.undigested) {
        const cached = this.stories.get(story);
        this.warming.set(story, cached && !changedSince.has(story) ? cached : null);
      }
    }
    for (const story of this.undigested) {
      if (!this.session.hasStory(story)) {
        this.release(this.stories.get(story)?.digests ?? []);
        this.stories.delete(story);
      } else {
        let cached = this.stories.get(story);
        if (!this.warming.has(story)) {
          this.warming.set(story, cached && !changedSince.has(story) ? cached : null);
        }
        const warming = this.warming.get(story);
        if (warming && !cached?.digests && (cached !== warming || changedSince.has(story))) {
          this.refresh();
          this.store(story, null, [this.session.storySegments(story)]);
          cached = this.stories.get(story);
        }
        if (warming && cached && !cached.digests) {
          const digests = this.session.storySegmentUnitDigests(story);
          const units = splitUnits(cached.segments);
          if (digests.length === units.length) {
            this.stories.delete(story);
            this.store(story, digests, units);
          }
        }
      }
      this.undigested.delete(story);
      this.warming.delete(story);
      if (
        deadline &&
        (performance.now() - start >= IDLE_SLICE_MS ||
          (deadline.timeRemaining() <= 1 && !deadline.didTimeout))
      ) {
        break;
      }
    }
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
