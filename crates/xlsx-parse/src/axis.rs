//! Source-to-current row and column addresses for a preserved sheet.

use std::ops::Range;

use xlsx_model::addr::{MAX_COLS, MAX_ROWS};

/// A monotone partial map from source index to current index on one axis.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AxisMap {
    segments: Vec<Segment>,
    limit: u32,
}

/// `len` consecutive source indices from `source` that now start at `current`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Segment {
    source: u32,
    len: u32,
    current: u32,
}

impl AxisMap {
    pub fn identity(limit: u32) -> Self {
        Self {
            segments: vec![Segment {
                source: 0,
                len: limit,
                current: 0,
            }],
            limit,
        }
    }

    pub fn is_identity(&self) -> bool {
        self.segments
            == [Segment {
                source: 0,
                len: self.limit,
                current: 0,
            }]
    }

    /// Current index of a source index; `None` when deleted or past the limit.
    pub fn current(&self, source: u32) -> Option<u32> {
        let index = self
            .segments
            .partition_point(|segment| segment.source + segment.len <= source);
        let segment = self.segments.get(index)?;
        (source >= segment.source).then(|| segment.current + (source - segment.source))
    }

    /// Current index of a source index, or of the first source index after it that survives
    /// when it was deleted: where a deleted row or column's neighbour now sits. `None` when no
    /// later index survives.
    pub fn current_at_or_after(&self, source: u32) -> Option<u32> {
        let index = self
            .segments
            .partition_point(|segment| segment.source + segment.len <= source);
        let segment = self.segments.get(index)?;
        Some(segment.current + source.saturating_sub(segment.source))
    }

    /// The source index of a current index; `None` for an inserted one.
    pub fn source(&self, current: u32) -> Option<u32> {
        // a move keeps the segments ordered by source, not by current, so the
        // current index is searched rather than binary-partitioned
        self.segments.iter().find_map(|segment| {
            (current >= segment.current && current < segment.current + segment.len)
                .then(|| segment.source + (current - segment.current))
        })
    }

    /// Current spans of a source span, in order; gaps stay split.
    pub fn current_ranges(&self, source: Range<u32>) -> Vec<Range<u32>> {
        Self::ranges(&self.segments, source)
    }

    pub(crate) fn current_ranges_bounded(&self, source: Range<u32>) -> Vec<Range<u32>> {
        let first = self
            .segments
            .partition_point(|segment| segment.source + segment.len <= source.start);
        let count = self.segments[first..].partition_point(|segment| segment.source < source.end);
        Self::ranges(&self.segments[first..first + count], source)
    }

    fn ranges(segments: &[Segment], source: Range<u32>) -> Vec<Range<u32>> {
        let mut ranges: Vec<Range<u32>> = Vec::new();
        let mut cursor = source.start;
        for segment in segments {
            let start = segment.source.max(source.start);
            let end = (segment.source + segment.len).min(source.end);
            if start >= end {
                continue;
            }
            let current = segment.current + (start - segment.source);
            let range = current..current + (end - start);
            match ranges.last_mut() {
                Some(last) if last.end == range.start && cursor == start => {
                    last.end = range.end;
                }
                _ => ranges.push(range),
            }
            cursor = end;
        }
        ranges
    }

    /// Records `count` indices inserted at current index `at`.
    pub fn insert(&mut self, at: u32, count: u32) {
        let mut segments = Vec::with_capacity(self.segments.len() + 1);
        for segment in &self.segments {
            if segment.current + segment.len <= at {
                self.push(&mut segments, *segment);
            } else if segment.current >= at {
                self.push(
                    &mut segments,
                    Segment {
                        current: segment.current.saturating_add(count),
                        ..*segment
                    },
                );
            } else {
                let head = at - segment.current;
                self.push(
                    &mut segments,
                    Segment {
                        len: head,
                        ..*segment
                    },
                );
                self.push(
                    &mut segments,
                    Segment {
                        source: segment.source + head,
                        len: segment.len - head,
                        current: at.saturating_add(count),
                    },
                );
            }
        }
        self.segments = segments;
    }

    /// Records `count` indices deleted from current index `at`.
    pub fn delete(&mut self, at: u32, count: u32) {
        let end = at.saturating_add(count);
        let mut segments = Vec::with_capacity(self.segments.len() + 1);
        for segment in &self.segments {
            let segment_end = segment.current + segment.len;
            if segment_end <= at {
                self.push(&mut segments, *segment);
                continue;
            }
            if segment.current >= end {
                self.push(
                    &mut segments,
                    Segment {
                        current: segment.current - count,
                        ..*segment
                    },
                );
                continue;
            }
            if segment.current < at {
                let head = at - segment.current;
                self.push(
                    &mut segments,
                    Segment {
                        len: head,
                        ..*segment
                    },
                );
            }
            if segment_end > end {
                let skipped = end - segment.current;
                self.push(
                    &mut segments,
                    Segment {
                        source: segment.source + skipped,
                        len: segment.len - skipped,
                        current: at,
                    },
                );
            }
        }
        self.segments = segments;
    }

    /// Records `count` indices moved from current index `at` to the gap at
    /// `before` (a permutation, not a shift, so the map is no longer monotone).
    /// Both indices are current coordinates, as every other edit takes them.
    pub fn move_block(&mut self, at: u32, count: u32, before: u32) {
        let end = at.saturating_add(count);
        if count == 0 || (before >= at && before <= end) {
            return;
        }
        let target = if before > end { before - count } else { before };
        let permuted = |current: u32| {
            if current >= at && current < end {
                target + (current - at)
            } else if before > end && current >= end && current < before {
                current - count
            } else if before < at && current >= before && current < at {
                current + count
            } else {
                current
            }
        };
        // split every segment against the three current intervals the move
        // distinguishes, then reorder the pieces by source: the map stays
        // searchable by source and every lookup finds its permuted position
        let bounds = [at, end, before];
        let mut pieces: Vec<(u32, u32, u32)> = Vec::new();
        for segment in &self.segments {
            let segment_start = segment.current;
            let segment_end = segment.current + segment.len;
            let mut cuts = vec![segment_start, segment_end];
            cuts.extend(
                bounds
                    .into_iter()
                    .filter(|b| *b > segment_start && *b < segment_end),
            );
            cuts.sort_unstable();
            cuts.dedup();
            for window in cuts.windows(2) {
                let (start, stop) = (window[0], window[1]);
                if start >= stop {
                    continue;
                }
                let source = segment.source + (start - segment.current);
                let current = permuted(start);
                let len = stop - start;
                if current < self.limit {
                    pieces.push((source, len.min(self.limit - current), current));
                }
            }
        }
        pieces.sort_unstable_by_key(|(source, _, _)| *source);
        let mut segments = Vec::with_capacity(pieces.len());
        for (source, len, current) in pieces {
            self.push(
                &mut segments,
                Segment {
                    source,
                    len,
                    current,
                },
            );
        }
        self.segments = segments;
    }

    /// Appends a segment clipped to the limit, merging gapless neighbors.
    fn push(&self, segments: &mut Vec<Segment>, mut segment: Segment) {
        if segment.current >= self.limit {
            return;
        }
        segment.len = segment.len.min(self.limit - segment.current);
        if segment.len == 0 {
            return;
        }
        if let Some(last) = segments.last_mut()
            && last.source + last.len == segment.source
            && last.current + last.len == segment.current
        {
            last.len += segment.len;
            return;
        }
        segments.push(segment);
    }
}

/// Both axes of one preserved sheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetAxes {
    pub rows: AxisMap,
    pub cols: AxisMap,
}

impl Default for SheetAxes {
    fn default() -> Self {
        Self {
            rows: AxisMap::identity(MAX_ROWS),
            cols: AxisMap::identity(MAX_COLS),
        }
    }
}

impl SheetAxes {
    pub fn is_identity(&self) -> bool {
        self.rows.is_identity() && self.cols.is_identity()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn insert_shifts_everything_at_and_after_the_point() {
        let mut map = AxisMap::identity(100);
        map.insert(3, 2);
        assert_eq!(map.current(2), Some(2));
        assert_eq!(map.current(3), Some(5));
        assert_eq!(map.current(99), None);
        assert_eq!(map.source(3), None);
        assert_eq!(map.source(4), None);
        assert_eq!(map.source(5), Some(3));
        assert!(!map.is_identity());
    }

    #[test]
    fn a_moved_block_permutes_both_axes_and_round_trips() {
        let mut map = AxisMap::identity(10);
        map.move_block(1, 2, 5);
        // 1,2 step over nothing: they land in the gap before 5, and 3,4 close up
        assert_eq!(map.current(0), Some(0));
        assert_eq!(map.current(1), Some(3));
        assert_eq!(map.current(2), Some(4));
        assert_eq!(map.current(3), Some(1));
        assert_eq!(map.current(4), Some(2));
        assert_eq!(map.current(5), Some(5));
        assert_eq!(map.source(3), Some(1));
        assert_eq!(map.source(1), Some(3));
        assert!(!map.is_identity());
    }

    #[test]
    fn a_block_moved_earlier_lands_before_the_tracks_it_passed() {
        let mut map = AxisMap::identity(10);
        map.move_block(6, 2, 1);
        assert_eq!(map.current(1), Some(3));
        assert_eq!(map.current(5), Some(7));
        assert_eq!(map.current(6), Some(1));
        assert_eq!(map.current(7), Some(2));
        assert_eq!(map.current(8), Some(8));
        assert_eq!(map.source(1), Some(6));
        assert_eq!(map.source(7), Some(5));
    }

    #[test]
    fn a_move_after_an_insert_permutes_current_positions_not_source_ones() {
        let mut map = AxisMap::identity(10);
        map.insert(1, 1);
        // B moved to column C by the insert; moving A after D puts it on D and
        // closes the columns between, so B lands back on B
        assert_eq!(map.current(1), Some(2));
        map.move_block(0, 1, 4);
        assert_eq!(map.current(0), Some(3));
        assert_eq!(map.current(1), Some(1));
        assert_eq!(map.current(2), Some(2));
        assert_eq!(map.source(3), Some(0));
        assert_eq!(map.source(1), Some(1));
    }

    #[test]
    fn a_deleted_index_resolves_to_the_next_survivor() {
        let mut map = AxisMap::identity(100);
        map.delete(3, 2);
        assert_eq!(map.current_at_or_after(2), Some(2));
        assert_eq!(map.current_at_or_after(3), Some(3));
        assert_eq!(map.current_at_or_after(4), Some(3));
        assert_eq!(map.current_at_or_after(5), Some(3));
        map.insert(3, 2);
        assert_eq!(map.current_at_or_after(4), Some(5));
        map.delete(90, 10);
        assert_eq!(map.current_at_or_after(97), None);
    }

    #[test]
    fn delete_drops_the_span_and_pulls_later_indices_back() {
        let mut map = AxisMap::identity(100);
        map.delete(3, 2);
        assert_eq!(map.current(2), Some(2));
        assert_eq!(map.current(3), None);
        assert_eq!(map.current(4), None);
        assert_eq!(map.current(5), Some(3));
        assert_eq!(map.source(3), Some(5));
        assert_eq!(map.current_ranges(0..10), vec![0..3, 3..8]);
    }

    #[test]
    fn inverse_edits_restore_identity() {
        let mut map = AxisMap::identity(100);
        map.insert(3, 2);
        map.delete(3, 2);
        for index in 0..98 {
            assert_eq!(map.current(index), Some(index));
            assert_eq!(map.source(index), Some(index));
        }
        assert_eq!(map.current(98), None);
        assert_eq!(map.current(99), None);
        map.delete(10, 4);
        map.insert(10, 4);
        assert_eq!(map.current(9), Some(9));
        assert_eq!(map.current(10), None);
        assert_eq!(map.current(14), Some(14));
        assert_eq!(map.source(12), None);
    }

    #[test]
    fn current_ranges_split_around_inserted_indices() {
        let mut map = AxisMap::identity(100);
        map.insert(4, 3);
        assert_eq!(map.current_ranges(0..10), vec![0..4, 7..13]);
        assert_eq!(map.current_ranges(20..20), Vec::<Range<u32>>::new());
    }

    #[test]
    fn bounded_current_ranges_match_full_scan_after_axis_edits() {
        let mut map = AxisMap::identity(100);
        let mut maps = vec![map.clone()];
        map.insert(4, 3);
        maps.push(map.clone());
        map.delete(20, 2);
        maps.push(map.clone());
        map.insert(20, 2);
        maps.push(map.clone());
        map.delete(4, 3);
        maps.push(map.clone());
        for map in maps {
            for start in 0..=105 {
                for end in 0..=105 {
                    assert_eq!(
                        map.current_ranges_bounded(start..end),
                        map.current_ranges(start..end),
                        "{start}..{end}, {map:?}"
                    );
                }
            }
        }
    }
}
