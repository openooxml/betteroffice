/** Coordinates selection paint with the latest completed layout sequence. */
export class LayoutSelectionGate {
  #stateSeq = 0;
  #renderSeq = 0;
  #layoutUpdating = false;
  #pendingRender = false;
  #renderCallbacks = new Set<() => void>();

  setStateSeq(sequence: number): void {
    this.#stateSeq = sequence;
  }
  incrementStateSeq(): number {
    return ++this.#stateSeq;
  }
  getStateSeq(): number {
    return this.#stateSeq;
  }
  getRenderSeq(): number {
    return this.#renderSeq;
  }
  onLayoutStart(): void {
    this.#layoutUpdating = true;
  }
  onLayoutComplete(sequence: number): void {
    this.#renderSeq = sequence;
    this.#layoutUpdating = false;
    if (this.#pendingRender && this.isSafeToRender()) {
      this.#pendingRender = false;
      this.#executeRender();
    }
  }
  isSafeToRender(): boolean {
    return !this.#layoutUpdating && this.#renderSeq >= this.#stateSeq;
  }
  requestRender(): void {
    if (this.isSafeToRender()) this.#executeRender();
    else this.#pendingRender = true;
  }
  onRender(callback: () => void): () => void {
    this.#renderCallbacks.add(callback);
    return () => this.#renderCallbacks.delete(callback);
  }
  reset(): void {
    this.#stateSeq = 0;
    this.#renderSeq = 0;
    this.#layoutUpdating = false;
    this.#pendingRender = false;
  }
  getDebugInfo(): {
    stateSeq: number;
    renderSeq: number;
    layoutUpdating: boolean;
    hasPendingRender: boolean;
    isSafe: boolean;
  } {
    return {
      stateSeq: this.#stateSeq,
      renderSeq: this.#renderSeq,
      layoutUpdating: this.#layoutUpdating,
      hasPendingRender: this.#pendingRender,
      isSafe: this.isSafeToRender(),
    };
  }
  #executeRender(): void {
    for (const callback of this.#renderCallbacks) {
      try {
        callback();
      } catch (error) {
        console.error('LayoutSelectionGate: render callback error', error);
      }
    }
  }
}
