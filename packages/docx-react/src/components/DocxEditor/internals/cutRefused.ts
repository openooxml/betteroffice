/** @internal A cut refused because the clipboard cannot carry these selected embed kinds. */
export class DocxCutRefusedError extends Error {
  readonly leftOut: readonly string[];

  constructor(leftOut: readonly string[]) {
    super(`Cut refused: the clipboard cannot carry ${leftOut.join(', ')}`);
    this.name = 'DocxCutRefusedError';
    this.leftOut = leftOut;
  }
}
