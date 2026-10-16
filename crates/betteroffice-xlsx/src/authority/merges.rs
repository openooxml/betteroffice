use super::*;

pub(super) const MAX_MERGE_EDITS_PER_SHEET: usize = 4_096;
pub(super) const MERGE_PREFIX: &str = "merge:";

pub(super) type MergeEdits = BTreeMap<String, (CellRange, bool)>;

fn merge_edit(key: &str, value: &Out) -> Result<(CellRange, bool), String> {
    let range_key = key
        .strip_prefix(MERGE_PREFIX)
        .ok_or_else(|| "invalid merge edit prefix".to_string())?;
    if key.len() > 64 {
        return Err("merge edit key is too long".to_string());
    }
    let range =
        CellRange::parse_a1(range_key).map_err(|_| format!("invalid merge edit key {key}"))?;
    if range.to_a1() != range_key
        || range_key.contains('$')
        || range.start == range.end
        || range.start.row >= MAX_ROWS
        || range.end.row >= MAX_ROWS
        || range.start.col >= MAX_COLS
        || range.end.col >= MAX_COLS
    {
        return Err(format!("noncanonical or out of bounds merge edit {key}"));
    }
    let Out::Any(Any::Bool(present)) = value else {
        return Err(format!("merge edit {key} is not a boolean"));
    };
    Ok((range, *present))
}

pub(super) fn insert_merge_edit(
    edits: &mut MergeEdits,
    key: &str,
    value: &Out,
) -> Result<(), String> {
    if edits.len() >= MAX_MERGE_EDITS_PER_SHEET {
        return Err("too many sheet merge edits".to_string());
    }
    let edit = merge_edit(key, value)?;
    edits.insert(edit.0.to_a1(), edit);
    Ok(())
}

pub(super) fn clear_merge_edits(sheet: &MapRef, txn: &mut TransactionMut<'_>) {
    for key in sheet
        .keys(txn)
        .filter(|key| key.starts_with(MERGE_PREFIX))
        .map(str::to_string)
        .collect::<Vec<_>>()
    {
        sheet.remove(txn, &key);
    }
}

pub(super) fn project_merges<T: ReadTxn>(
    baseline: Vec<CellRange>,
    sheet: &MapRef,
    txn: &T,
) -> Result<Vec<CellRange>, String> {
    let mut edits = MergeEdits::new();
    for (key, value) in sheet.iter(txn) {
        if key.starts_with(MERGE_PREFIX) {
            insert_merge_edit(&mut edits, key, &value)?;
        }
    }
    let mut projection = MergeProjection::new(edits);
    let mut ranges = Vec::new();
    for range in baseline {
        ranges.extend(projection.baseline(range)?);
    }
    while let Some(range) = projection.next_edit() {
        ranges.extend(range);
    }
    Ok(ranges)
}

pub(super) struct MergeProjection {
    edits: MergeEdits,
    changed: bool,
    occupied: MergeIndex,
    baseline_occupied: MergeIndex,
}

impl MergeProjection {
    pub(super) fn new(edits: MergeEdits) -> Self {
        Self {
            changed: !edits.is_empty(),
            edits,
            occupied: MergeIndex::default(),
            baseline_occupied: MergeIndex::default(),
        }
    }

    pub(super) fn baseline(&mut self, range: CellRange) -> Result<Option<CellRange>, String> {
        if range.start.row > range.end.row
            || range.start.col > range.end.col
            || range.end.row >= MAX_ROWS
            || range.end.col >= MAX_COLS
        {
            return Err("merge baseline is out of bounds".to_string());
        }
        if self.baseline_occupied.overlaps(1, 0, MAX_COLS - 1, &range) {
            return Err("merged cell ranges overlap".to_string());
        }
        self.baseline_occupied.insert(1, 0, MAX_COLS - 1, &range);
        if self
            .edits
            .remove(&range.to_a1())
            .is_some_and(|(_, present)| !present)
        {
            return Ok(None);
        }
        Ok(self.choose(range))
    }

    pub(super) fn next_edit(&mut self) -> Option<Option<CellRange>> {
        let (_, (range, present)) = self.edits.pop_first()?;
        Some(if present { self.choose(range) } else { None })
    }

    fn choose(&mut self, range: CellRange) -> Option<CellRange> {
        if !self.changed {
            return Some(range);
        }
        if self.occupied.overlaps(1, 0, MAX_COLS - 1, &range) {
            return None;
        }
        self.occupied.insert(1, 0, MAX_COLS - 1, &range);
        Some(range)
    }
}

#[derive(Default)]
struct MergeIndex {
    nodes: BTreeMap<u32, MergeNode>,
}

#[derive(Default)]
struct MergeNode {
    all: RowIntervals,
    cover: RowIntervals,
}

#[derive(Default)]
struct RowIntervals(BTreeMap<u32, u32>);

impl RowIntervals {
    fn overlaps(&self, start: u32, end: u32) -> bool {
        self.0
            .range(..=end)
            .next_back()
            .is_some_and(|(_, &last)| last >= start)
    }

    fn insert(&mut self, mut start: u32, mut end: u32) {
        if let Some((&first, &last)) = self.0.range(..=start).next_back()
            && last >= start
        {
            self.0.remove(&first);
            start = first;
            end = end.max(last);
        }
        while let Some((&first, &last)) = self.0.range(start..=end).next() {
            self.0.remove(&first);
            end = end.max(last);
        }
        self.0.insert(start, end);
    }
}

impl MergeIndex {
    fn overlaps(&self, node: u32, first: u32, last: u32, range: &CellRange) -> bool {
        let Some(entry) = self.nodes.get(&node) else {
            return false;
        };
        if range.start.col <= first && last <= range.end.col {
            return entry.all.overlaps(range.start.row, range.end.row);
        }
        if entry.cover.overlaps(range.start.row, range.end.row) {
            return true;
        }
        let middle = (first + last) / 2;
        (range.start.col <= middle && self.overlaps(node * 2, first, middle, range))
            || (range.end.col > middle && self.overlaps(node * 2 + 1, middle + 1, last, range))
    }

    fn insert(&mut self, node: u32, first: u32, last: u32, range: &CellRange) {
        let entry = self.nodes.entry(node).or_default();
        entry.all.insert(range.start.row, range.end.row);
        if range.start.col <= first && last <= range.end.col {
            entry.cover.insert(range.start.row, range.end.row);
            return;
        }
        let middle = (first + last) / 2;
        if range.start.col <= middle {
            self.insert(node * 2, first, middle, range);
        }
        if range.end.col > middle {
            self.insert(node * 2 + 1, middle + 1, last, range);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn anchored_merge_keys_are_refused() {
        for key in ["merge:$A$1:$B$2", "merge:$A1:B2", "merge:A1:B$2"] {
            for present in [false, true] {
                assert!(
                    merge_edit(key, &Out::Any(Any::Bool(present))).is_err(),
                    "{key}"
                );
            }
        }
    }

    #[test]
    fn single_cell_merge_keys_are_refused() {
        for present in [false, true] {
            assert!(merge_edit("merge:A1", &Out::Any(Any::Bool(present))).is_err());
        }
    }

    #[test]
    fn merge_interval_index_matches_rectangle_overlap() {
        let mut index = MergeIndex::default();
        let mut accepted: Vec<CellRange> = Vec::new();
        for number in 0..1_000 {
            let row = (number * 7919) % MAX_ROWS;
            let col = (number * 97) % MAX_COLS;
            let range = CellRange::new(
                CellRef::new(row, col),
                CellRef::new(
                    (row + number % 10000).min(MAX_ROWS - 1),
                    (col + number % 1000).min(MAX_COLS - 1),
                ),
            );
            let overlaps = accepted.iter().any(|other| other.overlaps(&range));
            assert_eq!(index.overlaps(1, 0, MAX_COLS - 1, &range), overlaps);
            if !overlaps {
                index.insert(1, 0, MAX_COLS - 1, &range);
                accepted.push(range);
            }
        }
        let full = CellRange::new(CellRef::new(0, 0), CellRef::new(MAX_ROWS - 1, MAX_COLS - 1));
        assert!(index.overlaps(1, 0, MAX_COLS - 1, &full));
        index.insert(1, 0, MAX_COLS - 1, &full);
        assert!(index.overlaps(
            1,
            0,
            MAX_COLS - 1,
            &CellRange::parse_a1("XFD1048576").unwrap()
        ));
    }
}
