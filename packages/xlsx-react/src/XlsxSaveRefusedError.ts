/** Why a synchronous {@link XlsxEditorApi.save} did not return bytes. */
export class XlsxSaveRefusedError extends Error {
  constructor(readonly code: 'input-pending' | 'input-failed') {
    super(
      code === 'input-pending'
        ? 'Accepted input is still being written; await commands.execute("save", null)'
        : 'A cell entry could not be written; correct or discard it first'
    );
    this.name = 'XlsxSaveRefusedError';
  }
}
