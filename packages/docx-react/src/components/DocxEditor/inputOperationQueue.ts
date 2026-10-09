const MAX_RECENT_FAILURES = 16;

export class InputOperationQueue {
  private pending: Promise<void> = Promise.resolve();
  private interactionEpoch = 0;
  private failures = 0;
  private recentFailures: Array<{ seq: number; error: unknown }> = [];
  private depth = 0;

  constructor(
    private readonly reportError: (error: unknown) => void,
    private readonly onPendingChange?: (pending: boolean) => void
  ) {}

  /** Admits accepted input; its failure is reported and fails only what was waiting for it. */
  enqueue(operation: () => void | Promise<void>): void {
    void this.admit(operation, true).catch(() => undefined);
  }

  /**
   * Admits an operation in input order and settles with its own outcome; `inputLost` is whether
   * input failed after the `since` checkpoint.
   */
  run<T>(operation: (inputLost: boolean) => T | Promise<T>, since = this.failures): Promise<T> {
    return this.admit(() => operation(this.failures !== since), false);
  }

  /** Checkpoint for {@link run} and {@link flush} requested before their admission. */
  failureCheckpoint(): number {
    return this.failures;
  }

  hasPending(): boolean {
    return this.depth > 0;
  }

  idle(): Promise<void> {
    return this.pending;
  }

  /** Waits for accepted operations and rejects if input failed after the `since` checkpoint. */
  flush(since = this.failures): Promise<void> {
    return this.pending.then(() => {
      if (this.failures !== since) throw this.firstFailureAfter(since);
    });
  }

  captureInteractionEpoch(): number {
    return this.interactionEpoch;
  }

  advanceInteractionEpoch(): void {
    this.interactionEpoch += 1;
  }

  isInteractionEpochCurrent(epoch: number): boolean {
    return this.interactionEpoch === epoch;
  }

  private admit<T>(operation: () => T | Promise<T>, input: boolean): Promise<T> {
    this.depth += 1;
    if (this.depth === 1) this.notify(true);
    const result = this.pending.then(operation);
    this.pending = result.then(
      () => this.settle(),
      (error) => {
        if (input) {
          this.failures += 1;
          this.recentFailures.push({ seq: this.failures, error });
          if (this.recentFailures.length > MAX_RECENT_FAILURES) this.recentFailures.shift();
          this.reportError(error);
        }
        this.settle();
      }
    );
    return result;
  }

  private firstFailureAfter(since: number): unknown {
    const recent = this.recentFailures;
    return (recent.find(({ seq }) => seq > since) ?? recent[recent.length - 1])?.error;
  }

  private settle(): void {
    this.depth -= 1;
    if (this.depth === 0) this.notify(false);
  }

  private notify(pending: boolean): void {
    try {
      this.onPendingChange?.(pending);
    } catch (error) {
      this.reportError(error);
    }
  }
}
