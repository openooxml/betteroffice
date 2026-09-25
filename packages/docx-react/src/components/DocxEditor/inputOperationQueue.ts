export class InputOperationQueue {
  private pending: Promise<void> = Promise.resolve();
  private interactionEpoch = 0;
  private failure: { error: unknown } | null = null;

  constructor(private readonly reportError: (error: unknown) => void) {}

  enqueue(operation: () => void | Promise<void>): void {
    const pending = this.pending.then(operation, operation);
    this.pending = pending.catch((error) => {
      this.failure ??= { error };
      this.reportError(error);
    });
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
}
