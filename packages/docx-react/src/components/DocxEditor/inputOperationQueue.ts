export class InputOperationQueue {
  private pending: Promise<void> = Promise.resolve();
  private interactionEpoch = 0;
  private failure: { error: unknown } | null = null;
  private depth = 0;

  constructor(
    private readonly reportError: (error: unknown) => void,
    private readonly onPendingChange?: (pending: boolean) => void
  ) {}

  /** Admits accepted input; its failure marks the queue as having lost input. */
  enqueue(operation: () => void | Promise<void>): void {
    void this.admit(operation, true).catch(() => undefined);
  }

  /** Admits an operation in input order and settles with its own outcome. */
  run<T>(operation: () => T | Promise<T>): Promise<T> {
    return this.admit(operation, false);
  }

  /** Whether an earlier input operation failed. */
  get failed(): boolean {
    return this.failure !== null;
  }

  hasPending(): boolean {
    return this.depth > 0;
  }

  idle(): Promise<void> {
    return this.pending;
  }

  /** Waits for accepted operations and rejects if this queue has lost input. */
  flush(): Promise<void> {
    return this.pending.then(() => {
      if (this.failure) throw this.failure.error;
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
          this.failure ??= { error };
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
