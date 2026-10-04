use std::sync::Arc;

use crate::SourceContainer;

pub struct SourceContainerBuilder {
    bytes: Arc<[u8]>,
    position: usize,
}

impl SourceContainerBuilder {
    pub fn new(length: usize) -> Result<Self, String> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(length)
            .map_err(|error| format!("cannot allocate source bytes: {error}"))?;
        bytes.resize(length, 0);
        Ok(Self {
            bytes: bytes.into(),
            position: 0,
        })
    }

    pub fn push(&mut self, bytes: &[u8]) -> Result<(), String> {
        let end = self
            .position
            .checked_add(bytes.len())
            .filter(|end| *end <= self.bytes.len())
            .ok_or_else(|| "source bytes exceed declared length".to_owned())?;
        Arc::get_mut(&mut self.bytes)
            .ok_or_else(|| "source bytes are already shared".to_owned())?[self.position..end]
            .copy_from_slice(bytes);
        self.position = end;
        Ok(())
    }

    pub fn finish(self) -> Result<SourceContainer, String> {
        if self.position != self.bytes.len() {
            return Err("source bytes are shorter than declared length".to_owned());
        }
        Ok(SourceContainer(self.bytes))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_bytes_roundtrip_is_exact() {
        let bytes: Vec<_> = (0..=255).cycle().take(10_003).collect();
        for size in [1, 7, 1_024, bytes.len()] {
            let mut builder = SourceContainerBuilder::new(bytes.len()).unwrap();
            let pointer = builder.bytes.as_ptr();
            for chunk in bytes.chunks(size) {
                builder.push(chunk).unwrap();
            }
            let source = builder.finish().unwrap();
            assert_eq!(source.as_bytes(), bytes);
            assert_eq!(source.as_bytes().as_ptr(), pointer);
            assert_eq!(source.clone().as_bytes().as_ptr(), pointer);
        }
        let mut empty = SourceContainerBuilder::new(0).unwrap();
        empty.push(&[]).unwrap();
        assert!(empty.finish().unwrap().as_bytes().is_empty());
    }

    #[test]
    fn source_bytes_reject_overflow_and_short_input() {
        let mut builder = SourceContainerBuilder::new(2).unwrap();
        assert!(builder.push(&[1, 2, 3]).is_err());
        builder.push(&[1, 2]).unwrap();
        assert!(builder.push(&[3]).is_err());
        assert_eq!(builder.finish().unwrap().as_bytes(), [1, 2]);
        let mut builder = SourceContainerBuilder::new(2).unwrap();
        builder.push(&[1]).unwrap();
        assert!(builder.finish().is_err());
        assert!(SourceContainerBuilder::new(usize::MAX).is_err());
    }
}
