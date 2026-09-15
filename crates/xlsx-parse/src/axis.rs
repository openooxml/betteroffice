//! Where a preserved sheet's source rows and columns sit after the row and
//! column edits made since it was read. The model carries the cells at their
//! current addresses; this carries which source address each came from.

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

    /// The current index of a source index; `None` once deleted or pushed past
    /// the axis limit.
    pub fn current(&self, source: u32) -> Option<u32> {
        let index = self
            .segments
            .partition_point(|segment| segment.source + segment.len <= source);
        let segment = self.segments.get(index)?;
        (source >= segment.source).then(|| segment.current + (source - segment.source))
    }

    /// The source index of a current index; `None` for an inserted one.
    pub fn source(&self, current: u32) -> Option<u32> {
        let index = self
            .segments
            .partition_point(|segment| segment.current + segment.len <= current);
        let segment = self.segments.get(index)?;
        (current >= segment.current).then(|| segment.source + (current - segment.current))
    }

    /// The current spans a source span survives as, in order. Spans split by a
    /// deleted or inserted gap stay split: merging them would claim a mapping
    /// the deleted indices do not have.
    pub fn current_ranges(&self, source: Range<u32>) -> Vec<Range<u32>> {
        let mut ranges: Vec<Range<u32>> = Vec::new();
        let mut cursor = source.start;
        for segment in &self.segments {
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

    /// Appends a segment clipped to the axis limit, merging it into the previous
    /// one when both axes continue without a gap.
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
}
