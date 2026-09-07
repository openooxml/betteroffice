export class SessionRevision {
  private revision = 0;
  private saved = 0;

  get dirty(): boolean {
    return this.revision !== this.saved;
  }
  change(): void {
    this.revision += 1;
  }
  checkpoint(): number {
    return this.revision;
  }
  didSave(checkpoint: number): void {
    this.saved = checkpoint;
  }
}
