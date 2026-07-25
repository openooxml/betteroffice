interface PendingDocumentLoad {
  generation: number;
  resolve: () => void;
}

export class DocumentLoadGeneration {
  private generation = 0;
  private completedGeneration: number | null = null;
  private pending: PendingDocumentLoad | null = null;

  begin(): number {
    this.generation += 1;
    this.completedGeneration = null;
    this.resolvePending();
    return this.generation;
  }

  isCurrent(generation: number): boolean {
    return this.generation === generation;
  }

  waitForCompletion(generation: number): Promise<void> {
    if (!this.isCurrent(generation) || this.completedGeneration === generation) {
      return Promise.resolve();
    }
    return new Promise((resolve) => {
      this.pending = { generation, resolve };
    });
  }

  complete(generation: number): boolean {
    if (!this.isCurrent(generation) || this.completedGeneration === generation) {
      return false;
    }
    this.completedGeneration = generation;
    if (this.pending?.generation === generation) {
      this.resolvePending();
    }
    return true;
  }

  invalidate(): void {
    this.generation += 1;
    this.completedGeneration = null;
    this.resolvePending();
  }

  private resolvePending(): void {
    const pending = this.pending;
    this.pending = null;
    pending?.resolve();
  }
}
