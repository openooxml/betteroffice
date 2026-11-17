// @betteroffice/docx declaration utils/readDocxFile.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Shared file-input → docx-buffer reader.
 *
 * React (DocxEditor.tsx `handleDocxFileChange`) and Vue
 * (DocxEditor.vue `handleDocxFileChange`) had byte-equivalent
 * `await file.arrayBuffer()` + `name.replace(/\.docx$/i, '')` boilerplate
 * around their hidden file inputs. This helper folds the common steps
 * into one place so the two adapters can never drift on filename
 * normalization or on the "reset input.value so re-picking the same
 * file fires `change` again" detail.
 *
 * Returns `null` when the user cancelled the picker (no file chosen).
 */
export interface ReadDocxFileResult {
    /** ArrayBuffer ready to feed into `loadBuffer` / `parseDocx`. */
    buffer: ArrayBuffer;
    /** File name with the trailing `.docx` extension stripped. */
    name: string;
}
/**
 * Read the first selected file out of an `<input type="file">` change
 * event, return its ArrayBuffer + a stem-form name. Always resets
 * `input.value` so re-selecting the same file in the same picker fires
 * the next `change` event.
 */
export declare function readDocxFileFromInput(event: Event): Promise<ReadDocxFileResult | null>;
