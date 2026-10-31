use crate::{Result, Workbook};

#[derive(Clone, Debug, Default)]
pub struct OpenOptions {}

impl Workbook {
    pub fn open_with_options(bytes: &[u8], _options: &OpenOptions) -> Result<Self> {
        Self::open(bytes)
    }
}
