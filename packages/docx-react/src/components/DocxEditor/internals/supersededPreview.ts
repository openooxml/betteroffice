/**
 * A newer layout reached the session, with another revision preview or from a
 * worker pass run again, or another load or session took its worker over; its
 * own pass shows it. A cancellation, never a document error.
 */
export class SupersededPreviewError extends Error {
  constructor() {
    super('Superseded by a newer layout, document load or worker owner');
    this.name = 'SupersededPreviewError';
  }
}
