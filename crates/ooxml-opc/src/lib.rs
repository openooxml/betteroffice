//! OPC (DOCX) container read/write, compiled to WASM.
//!
//! Replaces jszip at the unzip/rezip trust boundary. The decompression budget
//! (total inflated bytes + entry count) and path-traversal rejection are
//! enforced here by construction, so a malicious `.docx` cannot exhaust memory
//! or carry out-of-tree part names into the parser. The pure `*_parts` functions
//! hold the logic and are unit-tested natively; the `#[wasm_bindgen]` wrappers
//! only marshal to/from JS `{ path: Uint8Array }` objects.

use std::collections::HashSet;
use std::io::{Cursor, Read, Write};

#[cfg(feature = "wasm")]
use js_sys::{Object, Reflect, Uint8Array};
#[cfg(feature = "wasm")]
use wasm_bindgen::prelude::*;

mod sanitize;

pub use sanitize::{
    DocumentKind, DocumentKindError, detect_package_kind, sanitize_package,
    sanitize_package_for_format,
};

/// A well-formed document stays far under this; a decompression bomb blows past it.
const MAX_TOTAL_UNCOMPRESSED_BYTES: u64 = 512 * 1024 * 1024;

/// No legitimate package carries this many parts.
const MAX_ENTRY_COUNT: usize = 5000;

/// Reject absolute, drive-letter, and any `..` entry name (checked on both
/// separators, since producers may emit backslashes).
#[cfg(test)]
fn is_safe_entry_path(name: &str) -> bool {
    normalized_security_path(name).is_some()
}

/// Normalize only for security comparisons. Returned archive paths retain
/// their authored spelling; this key catches slash/case/dot aliases that could
/// otherwise name the same security-sensitive OPC part twice.
fn normalized_security_path(name: &str) -> Option<String> {
    let normalized = name.replace('\\', "/");
    if normalized.starts_with('/') {
        return None;
    }
    if normalized.len() >= 2 && normalized.as_bytes()[1] == b':' {
        return None;
    }
    let mut segments = Vec::new();
    for segment in normalized.split('/') {
        match segment {
            "" | "." => {}
            ".." => return None,
            segment => segments.push(segment),
        }
    }
    if segments.is_empty() {
        return None;
    }
    Some(segments.join("/").to_lowercase())
}

/// Extract every non-directory entry as `(path, bytes)`, enforcing the
/// decompression budget and path-traversal guard. Reading is bounded by the
/// remaining byte budget so a lying size header cannot force unbounded output.
pub fn unzip_parts(data: &[u8]) -> Result<Vec<(String, Vec<u8>)>, String> {
    unzip_parts_with_limits(data, MAX_TOTAL_UNCOMPRESSED_BYTES)
}

/// As [`unzip_parts`], with a caller-supplied expanded-data budget. The budget only
/// tightens: it is capped at the container ceiling, never raised above it.
pub fn unzip_parts_with_limits(
    data: &[u8],
    max_expanded_bytes: u64,
) -> Result<Vec<(String, Vec<u8>)>, String> {
    let budget = max_expanded_bytes.min(MAX_TOTAL_UNCOMPRESSED_BYTES);
    let mut archive =
        zip::ZipArchive::new(Cursor::new(data)).map_err(|e| format!("bad zip: {e}"))?;

    if archive.len() > MAX_ENTRY_COUNT {
        return Err(format!("zip entry count exceeds {MAX_ENTRY_COUNT}"));
    }

    let mut parts = Vec::with_capacity(archive.len());
    let mut total: u64 = 0;
    let mut seen_paths = HashSet::new();

    for i in 0..archive.len() {
        let mut entry = archive
            .by_index(i)
            .map_err(|e| format!("bad zip entry: {e}"))?;
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_string();
        let Some(security_path) = normalized_security_path(&name) else {
            return Err(format!("unsafe zip entry path: {name}"));
        };
        if !seen_paths.insert(security_path) {
            return Err(format!("duplicate normalized zip entry path: {name}"));
        }

        // read at most (budget - total) + 1 bytes: one over the limit proves a bomb
        let remaining = budget - total;
        let mut buf = Vec::new();
        entry
            .by_ref()
            .take(remaining + 1)
            .read_to_end(&mut buf)
            .map_err(|e| format!("read failed for {name}: {e}"))?;
        if buf.len() as u64 > remaining {
            return Err(format!("inflated size exceeds {budget} bytes"));
        }
        total += buf.len() as u64;
        parts.push((name, buf));
    }

    Ok(parts)
}

/// Write `(path, bytes)` entries into a deflated zip, in the given order.
pub fn rezip_parts(entries: &[(String, Vec<u8>)]) -> Result<Vec<u8>, String> {
    validate_parts(entries)?;

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        for (name, bytes) in entries {
            write_deflated(&mut writer, name, bytes)?;
        }
        writer.finish().map_err(|e| format!("finish: {e}"))?;
    }
    Ok(cursor.into_inner())
}

fn validate_parts(entries: &[(String, Vec<u8>)]) -> Result<(), String> {
    if entries.len() > MAX_ENTRY_COUNT {
        return Err(format!("zip entry count exceeds {MAX_ENTRY_COUNT}"));
    }
    let mut seen_paths = HashSet::new();
    let mut total = 0_u64;
    for (name, bytes) in entries {
        let Some(security_path) = normalized_security_path(name) else {
            return Err(format!("unsafe zip entry path: {name}"));
        };
        if !seen_paths.insert(security_path) {
            return Err(format!("duplicate normalized zip entry path: {name}"));
        }
        total = total
            .checked_add(bytes.len() as u64)
            .ok_or_else(|| format!("inflated size exceeds {MAX_TOTAL_UNCOMPRESSED_BYTES} bytes"))?;
        if total > MAX_TOTAL_UNCOMPRESSED_BYTES {
            return Err(format!(
                "inflated size exceeds {MAX_TOTAL_UNCOMPRESSED_BYTES} bytes"
            ));
        }
    }
    Ok(())
}

fn write_deflated(
    writer: &mut zip::ZipWriter<&mut Cursor<Vec<u8>>>,
    name: &str,
    bytes: &[u8],
) -> Result<(), String> {
    writer
        .start_file(
            name,
            zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated),
        )
        .map_err(|e| format!("start_file {name}: {e}"))?;
    writer
        .write_all(bytes)
        .map_err(|e| format!("write {name}: {e}"))
}

/// Source bytes retained so unchanged members re-emit verbatim on save; equality is always true.
#[derive(Clone, Default)]
pub struct SourceContainer(std::sync::Arc<[u8]>);

impl SourceContainer {
    pub fn new(bytes: Vec<u8>) -> Self {
        Self(bytes.into())
    }

    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }
}

impl std::fmt::Debug for SourceContainer {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SourceContainer")
            .field(&self.0.len())
            .finish()
    }
}

impl PartialEq for SourceContainer {
    fn eq(&self, _: &Self) -> bool {
        true
    }
}

impl Eq for SourceContainer {}

/// As [`rezip_parts`], but unchanged members copy the source's compressed
/// payload verbatim; unreadable `source` falls back to full rezip.
pub fn rezip_parts_preserving(
    entries: &[(String, Vec<u8>)],
    source: &[u8],
) -> Result<Vec<u8>, String> {
    validate_parts(entries)?;

    let Ok(mut archive) = zip::ZipArchive::new(Cursor::new(source)) else {
        return rezip_parts(entries);
    };
    if archive.len() > MAX_ENTRY_COUNT {
        return rezip_parts(entries);
    }

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        for (name, bytes) in entries {
            if !copy_unchanged_member(&mut archive, &mut writer, name, bytes)? {
                write_deflated(&mut writer, name, bytes)?;
            }
        }
        writer.finish().map_err(|e| format!("finish: {e}"))?;
    }
    Ok(cursor.into_inner())
}

/// Copy `name`'s compressed source member into `writer`; false means re-deflate.
fn copy_unchanged_member(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    writer: &mut zip::ZipWriter<&mut Cursor<Vec<u8>>>,
    name: &str,
    bytes: &[u8],
) -> Result<bool, String> {
    let Some(index) = archive.index_for_name(name) else {
        return Ok(false);
    };
    let Ok(mut file) = archive.by_index(index) else {
        return Ok(false);
    };
    if file.is_dir() || file.size() != bytes.len() as u64 || !inflated_matches(&mut file, bytes) {
        return Ok(false);
    }
    drop(file);
    let file = archive
        .by_index(index)
        .map_err(|e| format!("reopen {name}: {e}"))?;
    writer
        .raw_copy_file_rename(file, name)
        .map_err(|e| format!("copy {name}: {e}"))?;
    Ok(true)
}

fn inflated_matches(file: &mut zip::read::ZipFile<'_, Cursor<&[u8]>>, bytes: &[u8]) -> bool {
    let mut rest = bytes;
    let mut chunk = [0u8; 64 * 1024];
    loop {
        match file.read(&mut chunk) {
            Ok(0) => return rest.is_empty(),
            Ok(n) => {
                if rest.len() < n || rest[..n] != chunk[..n] {
                    return false;
                }
                rest = &rest[n..];
            }
            Err(_) => return false,
        }
    }
}

/// Unzip a DOCX; returns a JS object `{ [path]: Uint8Array }`.
#[cfg(feature = "wasm")]
#[wasm_bindgen]
pub fn unzip_docx(data: &[u8]) -> Result<JsValue, JsValue> {
    let parts = unzip_parts(data).map_err(|e| JsValue::from_str(&e))?;
    let out = Object::new();
    for (name, bytes) in parts {
        let arr = Uint8Array::from(bytes.as_slice());
        Reflect::set(&out, &JsValue::from_str(&name), &arr)?;
    }
    Ok(out.into())
}

/// Rezip from a JS object `{ [path]: Uint8Array }` into a DOCX byte array.
#[cfg(feature = "wasm")]
#[wasm_bindgen]
pub fn rezip_docx(entries: JsValue) -> Result<Vec<u8>, JsValue> {
    let obj: Object = entries
        .dyn_into()
        .map_err(|_| JsValue::from_str("rezip_docx: expected an object"))?;
    let mut collected: Vec<(String, Vec<u8>)> = Vec::new();
    let keys = Object::keys(&obj);
    for key in keys.iter() {
        let name = key
            .as_string()
            .ok_or_else(|| JsValue::from_str("rezip_docx: non-string key"))?;
        let value = Reflect::get(&obj, &key)?;
        let arr = Uint8Array::new(&value);
        collected.push((name, arr.to_vec()));
    }
    rezip_parts(&collected).map_err(|e| JsValue::from_str(&e))
}

#[cfg(feature = "wasm")]
#[wasm_bindgen(js_name = sanitizeOoxml)]
pub fn sanitize_ooxml(data: &[u8], expected_format: &str) -> Result<Vec<u8>, JsValue> {
    sanitize_package_for_format(data, expected_format).map_err(|error| JsValue::from_str(&error))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> Vec<(String, Vec<u8>)> {
        vec![
            ("[Content_Types].xml".into(), b"<Types/>".to_vec()),
            ("word/document.xml".into(), b"<w:document/>".to_vec()),
            ("word/media/image1.png".into(), vec![0x89, 0x50, 0x4e, 0x47]),
        ]
    }

    #[test]
    fn round_trips_parts() {
        let zipped = rezip_parts(&sample()).expect("rezip");
        let back = unzip_parts(&zipped).expect("unzip");
        assert_eq!(back, sample());
    }

    #[test]
    fn rejects_traversal_paths() {
        assert!(!is_safe_entry_path("../evil.xml"));
        assert!(!is_safe_entry_path("word/../../etc/passwd"));
        assert!(!is_safe_entry_path("/etc/passwd"));
        assert!(!is_safe_entry_path("C:/windows"));
        assert!(!is_safe_entry_path("word\\..\\..\\x"));
        assert!(is_safe_entry_path("word/document.xml"));
        assert!(is_safe_entry_path("word/my..file.xml"));
        assert!(!is_safe_entry_path("."));
    }

    #[test]
    fn rejects_traversal_on_read_and_write() {
        assert!(rezip_parts(&[("../escape.xml".into(), b"x".to_vec())]).is_err());

        // Bypass rezip_parts to exercise the independent read-side guard.
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            writer
                .start_file("../escape.xml", zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(b"x").unwrap();
            writer.finish().unwrap();
        }
        let zipped = cursor.into_inner();
        let err = unzip_parts(&zipped).unwrap_err();
        assert!(err.contains("unsafe"));
    }

    #[test]
    fn rejects_duplicate_normalized_paths_on_read_and_write() {
        let entries = vec![
            ("word/document.xml".into(), b"a".to_vec()),
            ("WORD//./document.xml".into(), b"b".to_vec()),
        ];
        assert!(rezip_parts(&entries).unwrap_err().contains("duplicate"));

        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            for (name, bytes) in &entries {
                writer
                    .start_file(name, zip::write::SimpleFileOptions::default())
                    .unwrap();
                writer.write_all(bytes).unwrap();
            }
            writer.finish().unwrap();
        }
        assert!(
            unzip_parts(&cursor.into_inner())
                .unwrap_err()
                .contains("duplicate")
        );
    }

    fn stored_zip(entries: &[(String, Vec<u8>)], stored: &str) -> Vec<u8> {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            for (name, bytes) in entries {
                let options = zip::write::SimpleFileOptions::default().compression_method(
                    if name == stored {
                        zip::CompressionMethod::Stored
                    } else {
                        zip::CompressionMethod::Deflated
                    },
                );
                writer.start_file(name, options).unwrap();
                writer.write_all(bytes).unwrap();
            }
            writer.finish().unwrap();
        }
        cursor.into_inner()
    }

    fn entry_method(zip_bytes: &[u8], name: &str) -> zip::CompressionMethod {
        let mut archive = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
        archive.by_name(name).unwrap().compression()
    }

    #[test]
    fn preserving_copies_verbatim_and_keeps_methods() {
        let entries = sample();
        let source = stored_zip(&entries, "word/media/image1.png");
        let out = rezip_parts_preserving(&entries, &source).expect("rezip");
        assert_eq!(unzip_parts(&out).unwrap(), entries);
        assert_eq!(
            entry_method(&out, "word/media/image1.png"),
            zip::CompressionMethod::Stored
        );
        assert_eq!(
            entry_method(&out, "word/document.xml"),
            zip::CompressionMethod::Deflated
        );
    }

    #[test]
    fn preserving_deflates_changed_and_new_parts() {
        let source = stored_zip(&sample(), "word/media/image1.png");
        let mut entries = sample();
        entries[1].1 = b"<w:document>edited</w:document>".to_vec();
        entries.push(("word/extra.xml".into(), b"<x/>".to_vec()));
        let out = rezip_parts_preserving(&entries, &source).expect("rezip");
        assert_eq!(unzip_parts(&out).unwrap(), entries);
        assert_eq!(
            entry_method(&out, "word/document.xml"),
            zip::CompressionMethod::Deflated
        );
    }

    #[test]
    fn preserving_rejects_same_size_forgery() {
        let source = stored_zip(&sample(), "word/media/image1.png");
        let mut entries = sample();
        // Same length, different content: must re-deflate, not copy.
        entries[2].1 = vec![0x89, 0x50, 0x4e, 0x00];
        let out = rezip_parts_preserving(&entries, &source).expect("rezip");
        assert_eq!(unzip_parts(&out).unwrap(), entries);
    }

    #[test]
    fn preserving_falls_back_when_source_is_not_a_zip() {
        let out = rezip_parts_preserving(&sample(), b"not a zip").expect("rezip");
        assert_eq!(unzip_parts(&out).unwrap(), sample());
    }

    #[test]
    fn preserving_still_enforces_entry_guards() {
        assert!(rezip_parts_preserving(&[("../x".into(), b"x".to_vec())], b"").is_err());
        let entries = vec![
            ("word/document.xml".into(), b"a".to_vec()),
            ("WORD//./document.xml".into(), b"b".to_vec()),
        ];
        assert!(
            rezip_parts_preserving(&entries, b"")
                .unwrap_err()
                .contains("duplicate")
        );
    }

    #[test]
    fn caller_budget_only_tightens_the_expanded_data_ceiling() {
        let mut entries = sample();
        entries.push(("word/media/filler.bin".into(), vec![0; 4 * 1024 * 1024]));
        let zipped = rezip_parts(&entries).expect("rezip");
        assert!(zipped.len() < 64 * 1024, "fixture must stay compressible");

        assert!(
            unzip_parts_with_limits(&zipped, 1024 * 1024)
                .unwrap_err()
                .contains("inflated size exceeds")
        );
        assert_eq!(
            unzip_parts_with_limits(&zipped, MAX_TOTAL_UNCOMPRESSED_BYTES).unwrap(),
            entries
        );
        assert_eq!(unzip_parts_with_limits(&zipped, u64::MAX).unwrap(), entries);
    }
}
