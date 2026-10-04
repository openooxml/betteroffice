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

/**
 * @internal
 * Dirty stories of an editor that keeps a main-thread projection cache and saves in the worker.
 * A worker save projects the stories the main-thread save would project at the same point.
 */
export class EditorDirtyStories {
  /** Stories changed since the main-thread projection cache last advanced. */
  readonly projection = new DirtyProjectionStories();
  private readonly workerSave = new DirtyProjectionStories();
  private projections = 0;
  private savedAtProjection = 0;
  private adoptingWorkerSaveUpdates = 0;

  add(story: string): void {
    if (this.adoptingWorkerSaveUpdates > 0) return;
    this.projection.add(story);
    this.workerSave.add(story);
  }

  /** Adopts a worker save's own updates without marking stories. */
  adoptWorkerSaveUpdates(apply: () => void): void {
    this.adoptingWorkerSaveUpdates++;
    try {
      apply();
    } finally {
      this.adoptingWorkerSaveUpdates--;
    }
  }

  /** The main-thread projection cache advanced over `projection`'s stories. */
  projected(): void {
    this.projection.clear();
    this.projections++;
  }

  captureWorkerSave(): { stories: string[]; clear(): void } {
    const captured = this.workerSave.capture();
    const projections = this.projections;
    const every = projections !== this.savedAtProjection && this.projection.projectionOptions() === undefined;
    return {
      stories: every ? [] : captured.stories,
      clear: () => {
        captured.clear();
        this.savedAtProjection = projections;
      },
    };
  }

  clear(): void {
    this.projection.clear();
    this.workerSave.clear();
    this.savedAtProjection = this.projections;
  }
}

/**
 * @internal
 * Runs one session's worker saves in call order. Each save captures its stories once the
 * previous save settled and clears them on success, as consecutive main-thread saves project.
 */
export function serialWorkerSaves(
  dirty: EditorDirtyStories
): <T>(save: (stories: string[]) => Promise<T>) => Promise<T> {
  let previous: Promise<unknown> = Promise.resolve();
  return <T>(save: (stories: string[]) => Promise<T>): Promise<T> => {
    const saving = previous.then(async () => {
      const captured = dirty.captureWorkerSave();
      const saved = await save(captured.stories);
      captured.clear();
      return saved;
    });
    previous = saving.catch(() => undefined);
    return saving;
  };
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
