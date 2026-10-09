use std::mem::{MaybeUninit, size_of};
use std::sync::Arc;

use crate::SourceContainer;

#[doc(hidden)]
pub struct SourceContainerBuilder {
    bytes: Arc<[u8]>,
    position: usize,
}

#[doc(hidden)]
pub struct SourceContainerInitializer {
    bytes: Arc<[MaybeUninit<u8>]>,
    position: usize,
}

impl SourceContainerInitializer {
    #[doc(hidden)]
    pub fn new(length: usize) -> Result<Self, String> {
        if length > isize::MAX as usize - 2 * size_of::<usize>() {
            return Err("cannot allocate source bytes: length overflows".to_owned());
        }
        Ok(Self {
            bytes: Arc::new_uninit_slice(length),
            position: 0,
        })
    }

    #[doc(hidden)]
    pub fn advance(&mut self, max_bytes: usize) -> Result<usize, String> {
        if max_bytes == 0 {
            return Err("source initialization budget must be positive".to_owned());
        }
        let count = max_bytes.min(self.bytes.len() - self.position);
        let bytes = Arc::get_mut(&mut self.bytes)
            .ok_or_else(|| "source bytes are already shared".to_owned())?;
        for byte in &mut bytes[self.position..self.position + count] {
            byte.write(0);
        }
        self.position += count;
        Ok(count)
    }

    #[doc(hidden)]
    pub fn is_ready(&self) -> bool {
        self.position == self.bytes.len()
    }

    #[doc(hidden)]
    pub fn finish(self) -> Result<SourceContainerBuilder, String> {
        if !self.is_ready() {
            return Err("source initialization is incomplete".to_owned());
        }
        Ok(SourceContainerBuilder {
            // Every element was initialized by advance before is_ready became true.
            bytes: unsafe { self.bytes.assume_init() },
            position: 0,
        })
    }
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
    fn source_bytes_incremental_initialization_is_bounded_and_exact() {
        let bytes: Vec<_> = (0..=255).cycle().take(10_003).collect();
        for budget in [1, 7, 256, bytes.len()] {
            let mut initializer = SourceContainerInitializer::new(bytes.len()).unwrap();
            let pointer = initializer.bytes.as_ptr().cast::<u8>();
            let mut initialized = 0;
            while !initializer.is_ready() {
                let count = initializer.advance(budget).unwrap();
                assert!(count <= budget);
                assert!(count > 0);
                initialized += count;
                assert_eq!(initializer.position, initialized);
            }
            assert_eq!(initialized, bytes.len());
            assert_eq!(initializer.advance(budget).unwrap(), 0);
            let mut builder = initializer.finish().unwrap();
            assert_eq!(builder.bytes.as_ptr(), pointer);
            for chunk in bytes.chunks(budget) {
                assert!(chunk.len() <= budget);
                builder.push(chunk).unwrap();
            }
            let source = builder.finish().unwrap();
            assert_eq!(source.as_bytes(), bytes);
            assert_eq!(source.as_bytes().as_ptr(), pointer);
        }
        let empty = SourceContainerInitializer::new(0).unwrap();
        assert!(empty.is_ready());
        assert!(
            empty
                .finish()
                .unwrap()
                .finish()
                .unwrap()
                .as_bytes()
                .is_empty()
        );
        assert!(SourceContainerInitializer::new(usize::MAX).is_err());
        let mut incomplete = SourceContainerInitializer::new(3).unwrap();
        assert!(incomplete.advance(0).is_err());
        assert_eq!(incomplete.advance(2).unwrap(), 2);
        assert!(incomplete.finish().is_err());
    }

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
