//! OPC (OOXML zip) container read/write shared across format cores, with
//! decompression budget and path-traversal rejection enforced by construction.

use std::io::{Cursor, Read, Write};

const MAX_TOTAL_UNCOMPRESSED_BYTES: u64 = 512 * 1024 * 1024;

const MAX_ENTRY_COUNT: usize = 5000;

/// Rejects absolute, drive-letter, and `..` entry names, on both separators.
pub fn is_safe_entry_path(name: &str) -> bool {
    let normalized = name.replace('\\', "/");
    if normalized.starts_with('/') {
        return false;
    }
    if normalized.len() >= 2 && normalized.as_bytes()[1] == b':' {
        return false;
    }
    !normalized.split('/').any(|segment| segment == "..")
}

/// Extracts every non-directory entry as `(path, bytes)`, enforcing the
/// decompression budget and traversal guard.
pub fn unzip_parts(data: &[u8]) -> Result<Vec<(String, Vec<u8>)>, String> {
    let mut archive =
        zip::ZipArchive::new(Cursor::new(data)).map_err(|e| format!("bad zip: {e}"))?;

    if archive.len() > MAX_ENTRY_COUNT {
        return Err(format!("zip entry count exceeds {MAX_ENTRY_COUNT}"));
    }

    let mut parts = Vec::with_capacity(archive.len());
    let mut total: u64 = 0;

    for i in 0..archive.len() {
        let mut entry = archive
            .by_index(i)
            .map_err(|e| format!("bad zip entry: {e}"))?;
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_string();
        if !is_safe_entry_path(&name) {
            return Err(format!("unsafe zip entry path: {name}"));
        }

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

/// Writes `(path, bytes)` entries into a deflated zip, in the given order;
/// entry names pass the same traversal guard as the read side.
pub fn rezip_parts(entries: &[(String, Vec<u8>)]) -> Result<Vec<u8>, String> {
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, bytes) in entries {
            if !is_safe_entry_path(name) {
                return Err(format!("unsafe zip entry path: {name}"));
            }
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

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> Vec<(String, Vec<u8>)> {
        vec![
            ("[Content_Types].xml".into(), b"<Types/>".to_vec()),
            ("word/document.xml".into(), b"<w:document/>".to_vec()),
            ("xl/workbook.xml".into(), b"<workbook/>".to_vec()),
            ("xl/media/image1.png".into(), vec![0x89, 0x50, 0x4e, 0x47]),
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
        assert!(!is_safe_entry_path("xl/../../etc/passwd"));
        assert!(!is_safe_entry_path("/etc/passwd"));
        assert!(!is_safe_entry_path("C:/windows"));
        assert!(!is_safe_entry_path("word\\..\\..\\x"));
        assert!(is_safe_entry_path("word/document.xml"));
        assert!(is_safe_entry_path("xl/workbook.xml"));
        assert!(is_safe_entry_path("xl/my..file.xml"));
    }

    #[test]
    fn rejects_traversal_on_rezip() {
        let err = rezip_parts(&[("../escape.xml".into(), b"x".to_vec())]).unwrap_err();
        assert!(err.contains("unsafe"));
    }

    #[test]
    fn rejects_traversal_on_unzip() {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            writer.start_file("../escape.xml", options).unwrap();
            writer.write_all(b"x").unwrap();
            writer.finish().unwrap();
        }
        let err = unzip_parts(&cursor.into_inner()).unwrap_err();
        assert!(err.contains("unsafe"));
    }
}
