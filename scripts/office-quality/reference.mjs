export const MAX_REFERENCE_PAGES = 250;

export function validateReferenceMetadata(metadata, id) {
  const reference = metadata?.reference;
  const maximumPages = metadata?.format === 'docx' ? MAX_REFERENCE_PAGES : 100;
  if (
    reference?.status !== 'ok' ||
    reference.dpi !== 150 ||
    typeof reference.sha256 !== 'string' ||
    !/^[a-f0-9]{64}$/.test(reference.sha256) ||
    reference.sha256 !== metadata.source?.sha256 ||
    !Number.isInteger(reference.pages) ||
    reference.pages < 1 ||
    reference.pages > maximumPages ||
    !Array.isArray(metadata.reference_pages) ||
    reference.pages !== metadata.reference_pages.length
  )
    throw new Error(`Invalid Office reference: ${id}`);
}
