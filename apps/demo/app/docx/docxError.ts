import type { DocxErrorInfo } from "@betteroffice/docx-react";

export function isFatalDocxError(info?: DocxErrorInfo): boolean {
  return !info?.recoverable;
}
