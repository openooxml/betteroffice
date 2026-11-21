use crate::{Document, Result};

#[derive(Clone, Debug, Default)]
pub struct OpenOptions {}

impl Document {
    pub fn open_with_options(bytes: &[u8], _options: &OpenOptions) -> Result<Self> {
        Self::open(bytes)
    }
}
