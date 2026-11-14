//! ISO 19005-2 string-size limits shared by metadata, links and text.

/// ISO 19005-2 6.1.13: no PDF string may exceed 32767 bytes.
pub(crate) const MAX_PDF_STRING_BYTES: usize = 32767;

/// Error text for a PDF/A string that exceeds the ISO 19005-2 limit.
pub(crate) const PDFA_STRING_TOO_LONG: &str = "PDF/A-2u string exceeds 32767 bytes";

/// Encoded size of a `pdf_writer::TextStr` value: ASCII stays literal;
/// anything else becomes a UTF-16BE hex string with a BOM.
pub(crate) fn text_string_bytes(value: &str) -> usize {
    if value.bytes().all(|byte| matches!(byte, 32..=126)) {
        value.len()
    } else {
        2 + value.encode_utf16().count() * 2
    }
}

/// Whether a text string would exceed the ISO 19005-2 limit.
pub(crate) fn text_too_long(value: &str) -> bool {
    text_string_bytes(value) > MAX_PDF_STRING_BYTES
}

/// Whether a raw PDF string would exceed the ISO 19005-2 limit.
pub(crate) fn bytes_too_long(value: &[u8]) -> bool {
    value.len() > MAX_PDF_STRING_BYTES
}
