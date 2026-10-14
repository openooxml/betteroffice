use std::collections::BTreeMap;

use crate::{Lookup, ResolvedShape};

pub(crate) fn row_cells(
    shape: &ResolvedShape,
    section: &str,
    index: u32,
) -> BTreeMap<String, Lookup> {
    shape
        .sections
        .get(section)
        .and_then(|section| section.rows.get(&format!("IX:{index}")))
        .map(|row| row.cells.clone())
        .unwrap_or_default()
}
