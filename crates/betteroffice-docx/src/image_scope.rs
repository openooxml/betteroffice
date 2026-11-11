//! Part-scoped image lookup for PDF export tests.
//!
//! DOCX relationship ids are relative to the owning part. The PDF writer
//! resolves them with the same `\u{1f}`-separated keys the raster registry
//! stores; this module keeps that key construction in one place.

use std::collections::HashMap;

const SCOPE_SEPARATOR: char = '\u{1f}';

fn body_key(rel_id: &str) -> String {
    format!("{SCOPE_SEPARATOR}{rel_id}")
}

fn header_footer_key(region_rel_id: &str, rel_id: &str) -> String {
    format!("hf:{region_rel_id}{SCOPE_SEPARATOR}{rel_id}")
}

fn notes_key(kind: Option<&str>, rel_id: &str) -> String {
    let part = if kind == Some("endnote") {
        "endnotes"
    } else {
        "footnotes"
    };
    format!("{part}{SCOPE_SEPARATOR}{rel_id}")
}

/// A borrowed view over part-scoped image bytes.
pub(crate) struct ScopedImages<'a> {
    entries: &'a HashMap<String, Vec<u8>>,
}

impl<'a> ScopedImages<'a> {
    pub(crate) fn from_registry(entries: &'a HashMap<String, Vec<u8>>) -> Self {
        Self { entries }
    }

    pub(crate) fn bytes(&self, part: Option<&str>, rel_id: &str) -> Option<&'a [u8]> {
        let key = match part {
            None => body_key(rel_id),
            Some("footnotes") => notes_key(Some("footnote"), rel_id),
            Some("endnotes") => notes_key(Some("endnote"), rel_id),
            Some(region_rel_id) => header_footer_key(region_rel_id, rel_id),
        };
        self.entries
            .get(&key)
            .or_else(|| part.is_none().then(|| self.entries.get(rel_id)).flatten())
            .map(Vec::as_slice)
    }

    pub(crate) fn body(&self, rel_id: &str) -> Option<&'a [u8]> {
        self.entries
            .get(&body_key(rel_id))
            .or_else(|| self.entries.get(rel_id))
            .map(Vec::as_slice)
    }

    pub(crate) fn header_footer(&self, region_rel_id: &str, rel_id: &str) -> Option<&'a [u8]> {
        self.entries
            .get(&header_footer_key(region_rel_id, rel_id))
            .map(Vec::as_slice)
    }

    pub(crate) fn notes(&self, kind: Option<&str>, rel_id: &str) -> Option<&'a [u8]> {
        self.entries
            .get(&notes_key(kind, rel_id))
            .map(Vec::as_slice)
    }
}
