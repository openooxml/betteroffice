/**
 * What a failed open says about the file, read from the engine's message: its
 * open path reports plain strings, not codes. `bindings/docx-native/engine.test.ts`
 * opens real broken packages on both engines, so a changed message fails there.
 * `null` is anything else, such as a size limit or a crash, which says nothing
 * about the file; a limit can arrive wrapped in a damaged-looking prefix.
 */
export function openFailureKind(message: string): 'damaged' | 'unsupported' | null {
  if (LIMIT.test(message)) return null;
  if (DAMAGED_FILE.test(message)) return 'damaged';
  return UNSUPPORTED_CONTENT.test(message) ? 'unsupported' : null;
}

const LIMIT = /\blimit exceeded\b|\bexceeds\b|\btoo many\b/;
const DAMAGED_FILE =
  /\b(?:bad zip|not a Word document|unsafe zip entry path|duplicate normalized zip entry path|read failed for|inflates to other than its declared|unexpected EOF in|unexpected closing element in|malformed XML in|invalid XML in|invalid relationship in|(?:missing|invalid) \[Content_Types\]|(?:missing|invalid) _rels\/\.rels|missing main document|conflicting main document|missing recognized main document)/;
const UNSUPPORTED_CONTENT = /\b(?:unsupported geometry projection|missing (?:cell|control) (?:projection|paragraph))\b/;
