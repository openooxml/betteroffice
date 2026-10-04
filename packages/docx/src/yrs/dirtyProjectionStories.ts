import type { DocxProposalResult } from './proposals';
import type { YrsToDocumentOptions } from './yrsToDocument';

/** @internal */
export function dirtyProjectionStory(activeStory: string): string {
  return ['hf:', 'fn:', 'en:'].some((prefix) => activeStory.startsWith(prefix))
    ? activeStory.split(':', 2).join(':')
    : 'body';
}

/** @internal */
export class DirtyProjectionStories {
  private readonly stories = new Map<string, number>();
  private nextMark = 0;

  add(story: string): void {
    this.stories.set(dirtyProjectionStory(story), ++this.nextMark);
  }

  projectionOptions(): YrsToDocumentOptions | undefined {
    return this.stories.size > 0 ? { storyIds: new Set(this.stories.keys()) } : undefined;
  }

  capture(): { stories: string[]; clear(): void } {
    const captured = new Map(this.stories);
    return {
      stories: [...captured.keys()],
      clear: () => {
        for (const [story, mark] of captured) {
          if (this.stories.get(story) === mark) this.stories.delete(story);
        }
      },
    };
  }

  clear(): void {
    this.stories.clear();
  }
}

/** @internal */
export function proposalProjectionStories(
  known: ReadonlySet<string>,
  result: Extract<DocxProposalResult, { ok: true }>,
  changedStories: readonly string[]
): Set<string> {
  return new Set([
    ...result.snapshot.proposals
      .filter((proposal) => proposal.changed && !known.has(proposal.id))
      .map((proposal) => proposal.paragraph.story),
    ...changedStories,
  ]);
}
