//! Source `cellXfs` indices kept across a save when the model resolves them to
//! the same format as the index it carries now.

use std::cell::RefCell;
use std::collections::HashMap;

use xlsx_model::styles::{Stylesheet, Xf};
use xlsx_model::{Cell, CellRef, Sheet};

use crate::axis::SheetAxes;

/// Style equivalence between the source stylesheet and the one being written.
pub(crate) struct StyleMatch<'a> {
    pub(super) original: &'a Stylesheet,
    current: &'a Stylesheet,
    pairs: RefCell<HashMap<(u32, u32), bool>>,
    pub(super) derived: HashMap<u32, u32>,
}

impl<'a> StyleMatch<'a> {
    pub(crate) fn new(original: &'a Stylesheet, current: &'a Stylesheet) -> Self {
        Self {
            original,
            current,
            pairs: RefCell::new(HashMap::new()),
            derived: HashMap::new(),
        }
    }

    pub(crate) fn with_derived(mut self, derived: HashMap<u32, u32>) -> Self {
        self.derived = derived;
        self
    }

    /// Whether a source cell's index still means the format the model's index
    /// resolves to, in both the source and the written stylesheet.
    pub(crate) fn equivalent(&self, source: Option<u32>, current: Option<u32>) -> bool {
        if source == current {
            return true;
        }
        let (Some(source), Some(current)) = (source, current) else {
            return false;
        };
        *self
            .pairs
            .borrow_mut()
            .entry((source, current))
            .or_insert_with(|| self.resolve(source, current))
    }

    fn resolve(&self, source: u32, current: u32) -> bool {
        let in_range =
            |stylesheet: &Stylesheet, index: u32| (index as usize) < stylesheet.cell_xfs.len();
        if self.current.has_style_identity() && in_range(self.original, current) {
            return false;
        }
        if !in_range(self.original, source)
            || !in_range(self.current, source)
            || !in_range(self.current, current)
        {
            return false;
        }
        if current as usize >= self.original.cell_xfs.len()
            && self.current.resolved_format(Some(current)) == self.current.resolved_format(None)
        {
            return false;
        }
        let meant = self.original.resolved_format(Some(source));
        self.original.cell_xfs[source as usize] == self.current.cell_xfs[source as usize]
            && meant == self.current.resolved_format(Some(current))
            && meant == self.current.resolved_format(Some(source))
    }

    /// The index to write for a model cell whose source cell carried `source`.
    pub(crate) fn written(&self, source: Option<u32>, current: Option<u32>) -> Option<u32> {
        if let Some(index) = self.derived.get(&current.unwrap_or(u32::MAX)) {
            return Some(*index);
        }
        if self.equivalent(source, current) {
            source
        } else {
            current
        }
    }

    /// Cell equality with the style index compared by the format it resolves to.
    pub(crate) fn same_cell(&self, source: &Cell, current: &Cell) -> bool {
        source.value == current.value
            && source.formula == current.formula
            && self.equivalent(source.style, current.style)
    }

    pub(crate) fn same(&self, source: Option<&Cell>, current: Option<&Cell>) -> bool {
        match (source, current) {
            (Some(source), Some(current)) => self.same_cell(source, current),
            (None, None) => true,
            _ => false,
        }
    }
}

/// Where a regenerated cell finds the source cell whose index it may keep.
#[derive(Clone, Copy)]
pub(crate) struct SourceStyles<'a> {
    pub(crate) original: &'a Sheet,
    pub(crate) axes: &'a SheetAxes,
    pub(crate) styles: &'a StyleMatch<'a>,
}

impl SourceStyles<'_> {
    pub(crate) fn style(&self, at: CellRef, cell: &Cell) -> Option<u32> {
        let source = self
            .axes
            .rows
            .source(at.row)
            .zip(self.axes.cols.source(at.col))
            .map(|(row, col)| CellRef::new(row, col));
        match source.and_then(|source| self.original.cell(source)) {
            Some(original) => self.styles.written(original.style, cell.style),
            None if source.is_some() => self.styles.written(None, cell.style),
            None => self.styles.written(None, cell.style),
        }
    }
}

/// A restored entry read from a snapshot record written before format facets
/// keeps its source entry when it equals that entry's released reading; any
/// other entry keeps it only when equal.
pub fn keeps_source<T: PartialEq>(
    current: &T,
    source: &T,
    released: bool,
    released_reading: impl FnOnce(&T) -> T,
) -> bool {
    current == source || (released && *current == released_reading(source))
}

/// Whether `cell_xfs[index]` of `current` writes as the source stylesheet's.
pub(crate) fn keeps_source_xf(current: &Stylesheet, source: &Stylesheet, index: usize) -> bool {
    let released = current.is_released_xf(index);
    match (current.cell_xfs.get(index), source.cell_xfs.get(index)) {
        (Some(xf), Some(original)) => keeps_source(xf, original, released, Xf::released_reading),
        _ => false,
    }
}

/// Whether `current` writes back as the source stylesheet.
pub(crate) fn retains_source_styles(current: &Stylesheet, source: &Stylesheet) -> bool {
    if current == source {
        return true;
    }
    let count = current.cell_xfs.len();
    if count != source.cell_xfs.len()
        || !(0..count).any(|index| current.is_released_xf(index))
        || !(0..count).all(|index| keeps_source_xf(current, source, index))
    {
        return false;
    }
    let mut kept = current.clone();
    kept.cell_xfs.clone_from(&source.cell_xfs);
    kept == *source
}
