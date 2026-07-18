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

use js_sys::{Object, Reflect, Uint8Array};
use wasm_bindgen::prelude::*;

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
        let remaining = MAX_TOTAL_UNCOMPRESSED_BYTES - total;
        let mut buf = Vec::new();
        entry
            .by_ref()
            .take(remaining + 1)
            .read_to_end(&mut buf)
            .map_err(|e| format!("read failed for {name}: {e}"))?;
        if buf.len() as u64 > remaining {
            return Err(format!(
                "inflated size exceeds {MAX_TOTAL_UNCOMPRESSED_BYTES} bytes"
            ));
        }
        total += buf.len() as u64;
        parts.push((name, buf));
    }

    Ok(parts)
}

/// Write `(path, bytes)` entries into a deflated zip, in the given order.
pub fn rezip_parts(entries: &[(String, Vec<u8>)]) -> Result<Vec<u8>, String> {
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

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, bytes) in entries {
            writer
                .start_file(name, options)
                .map_err(|e| format!("start_file {name}: {e}"))?;
            writer
                .write_all(bytes)
                .map_err(|e| format!("write {name}: {e}"))?;
        }
        writer.finish().map_err(|e| format!("finish: {e}"))?;
    }
    Ok(cursor.into_inner())
}

/// Unzip a DOCX; returns a JS object `{ [path]: Uint8Array }`.
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
    rezip_parts(&collected)
        .map(|bytes| bytes)
        .map_err(|e| JsValue::from_str(&e))
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
}
