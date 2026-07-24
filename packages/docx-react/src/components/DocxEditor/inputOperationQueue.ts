export class InputOperationQueue {
  private pending: Promise<void> = Promise.resolve();

  constructor(private readonly reportError: (error: unknown) => void) {}

  enqueue(operation: () => void | Promise<void>): void {
    const pending = this.pending.then(operation, operation);
    this.pending = pending.catch(this.reportError);
  }

  idle(): Promise<void> {
    return this.pending;
  }
}
