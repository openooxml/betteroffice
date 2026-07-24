export class InputOperationQueue {
  private pending: Promise<void> = Promise.resolve();
  private interactionEpoch = 0;

  constructor(private readonly reportError: (error: unknown) => void) {}

  enqueue(operation: () => void | Promise<void>): void {
    const pending = this.pending.then(operation, operation);
    this.pending = pending.catch(this.reportError);
  }

  idle(): Promise<void> {
    return this.pending;
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
