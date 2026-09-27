export class InputOperationQueue {
  private pending: Promise<void> = Promise.resolve();
  private interactionEpoch = 0;
  private failures = 0;
  private lastFailure: unknown;
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
   * input accepted before it failed.
   */
  run<T>(operation: (inputLost: boolean) => T | Promise<T>): Promise<T> {
    const failures = this.failures;
    return this.admit(() => operation(this.failures !== failures), false);
  }

  hasPending(): boolean {
    return this.depth > 0;
  }

  idle(): Promise<void> {
    return this.pending;
  }

  /** Waits for accepted operations and rejects if one of them failed. */
  flush(): Promise<void> {
    const failures = this.failures;
    return this.pending.then(() => {
      if (this.failures !== failures) throw this.lastFailure;
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
          this.lastFailure = error;
          this.reportError(error);
        }
        this.settle();
      }
    );
    return result;
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
