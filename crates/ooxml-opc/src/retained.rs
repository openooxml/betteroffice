use std::collections::HashSet;
use std::io::{Cursor, Read};
use std::sync::Arc;

use crate::{MAX_ENTRY_COUNT, MAX_TOTAL_UNCOMPRESSED_BYTES, normalized_security_path};

/// The most a read reserves before the part's bytes arrive.
const PREALLOCATED: u64 = 32 * 1024 * 1024;

/// A package kept whole, its parts inflated one at a time when read. Opening
/// it checks every entry name as [`crate::unzip_parts`] does. A read inflates
/// at most the part's declared size, so a part whose size header lies fails
/// that read.
#[derive(Clone)]
pub struct RetainedPackage {
    archive: zip::ZipArchive<Cursor<Arc<[u8]>>>,
    parts: Arc<[PackagePart]>,
}

#[derive(Clone, Debug)]
struct PackagePart {
    name: String,
    index: usize,
    size: u64,
}

impl RetainedPackage {
    pub fn new(bytes: Arc<[u8]>) -> Result<Self, String> {
        let mut archive =
            zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| format!("bad zip: {e}"))?;
        if archive.len() > MAX_ENTRY_COUNT {
            return Err(format!("zip entry count exceeds {MAX_ENTRY_COUNT}"));
        }
        let mut parts = Vec::with_capacity(archive.len());
        let mut seen_paths = HashSet::new();
        for index in 0..archive.len() {
            let entry = archive
                .by_index_raw(index)
                .map_err(|e| format!("bad zip entry: {e}"))?;
            if entry.is_dir() {
                continue;
            }
            let name = entry.name().to_owned();
            let Some(security_path) = normalized_security_path(&name) else {
                return Err(format!("unsafe zip entry path: {name}"));
            };
            if !seen_paths.insert(security_path) {
                return Err(format!("duplicate normalized zip entry path: {name}"));
            }
            parts.push(PackagePart {
                name,
                index,
                size: entry.size(),
            });
        }
        Ok(Self {
            archive,
            parts: parts.into(),
        })
    }

    /// The name and declared inflated size of every part, in archive order.
    pub fn parts(&self) -> impl ExactSizeIterator<Item = (&str, u64)> {
        self.parts
            .iter()
            .map(|part| (part.name.as_str(), part.size))
    }

    /// The inflated bytes of the part at `position` in [`RetainedPackage::parts`].
    pub fn read(&self, position: usize) -> Result<Vec<u8>, String> {
        let part = self.part(position)?;
        let mut bytes =
            Vec::with_capacity(usize::try_from(part.size.min(PREALLOCATED)).unwrap_or(0));
        self.inflate(part, part.size + 1, &mut bytes)?;
        if bytes.len() as u64 != part.size {
            return Err(format!(
                "{} inflates to other than its declared {} bytes",
                part.name, part.size
            ));
        }
        Ok(bytes)
    }

    /// At most the first `len` inflated bytes of the part at `position`.
    pub fn read_prefix(&self, position: usize, len: usize) -> Result<Vec<u8>, String> {
        let part = self.part(position)?;
        let limit = part.size.min(len as u64);
        let mut bytes = Vec::with_capacity(usize::try_from(limit.min(PREALLOCATED)).unwrap_or(0));
        self.inflate(part, limit, &mut bytes)?;
        Ok(bytes)
    }

    fn part(&self, position: usize) -> Result<&PackagePart, String> {
        let part = self
            .parts
            .get(position)
            .ok_or_else(|| format!("no package part at {position}"))?;
        if part.size > MAX_TOTAL_UNCOMPRESSED_BYTES {
            return Err(format!(
                "inflated size exceeds {MAX_TOTAL_UNCOMPRESSED_BYTES} bytes"
            ));
        }
        Ok(part)
    }

    fn inflate(&self, part: &PackagePart, limit: u64, out: &mut Vec<u8>) -> Result<(), String> {
        let mut archive = self.archive.clone();
        let entry = archive
            .by_index(part.index)
            .map_err(|e| format!("bad zip entry {}: {e}", part.name))?;
        entry
            .take(limit)
            .read_to_end(out)
            .map_err(|e| format!("read failed for {}: {e}", part.name))?;
        Ok(())
    }
}

impl std::fmt::Debug for RetainedPackage {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("RetainedPackage")
            .field("parts", &self.parts.len())
            .finish()
    }
}

#[cfg(test)]
mod tests {
    use std::io::Write;

    use super::*;
    use crate::{rezip_parts, unzip_parts};

    fn package(entries: &[(&str, &[u8])]) -> Arc<[u8]> {
        let entries: Vec<(String, Vec<u8>)> = entries
            .iter()
            .map(|(name, bytes)| ((*name).to_owned(), bytes.to_vec()))
            .collect();
        rezip_parts(&entries).unwrap().into()
    }

    #[test]
    fn reads_each_part_as_unzip_inflates_it() {
        let bytes = package(&[
            ("word/document.xml", b"<w:document/>"),
            ("word/media/image1.png", &[0x89, 0x50, 0x4e, 0x47, 1, 2, 3]),
        ]);
        let retained = RetainedPackage::new(Arc::clone(&bytes)).unwrap();
        let eager = unzip_parts(&bytes).unwrap();
        assert_eq!(
            retained.parts().map(|(name, _)| name).collect::<Vec<_>>(),
            ["word/document.xml", "word/media/image1.png"]
        );
        for (position, (name, data)) in eager.iter().enumerate() {
            assert_eq!(
                retained.parts().nth(position).unwrap(),
                (name.as_str(), data.len() as u64)
            );
            assert_eq!(&retained.read(position).unwrap(), data);
        }
        assert_eq!(
            retained.read_prefix(1, 4).unwrap(),
            [0x89, 0x50, 0x4e, 0x47]
        );
        assert_eq!(retained.read_prefix(0, 1024).unwrap(), b"<w:document/>");
        assert!(retained.read(2).is_err());
    }

    #[test]
    fn refuses_what_unzip_refuses() {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            writer
                .start_file("../escape.xml", zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(b"x").unwrap();
            writer.finish().unwrap();
        }
        let bytes: Arc<[u8]> = cursor.into_inner().into();
        assert!(RetainedPackage::new(bytes).unwrap_err().contains("unsafe"));
    }
}
