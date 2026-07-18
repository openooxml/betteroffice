//! Paragraph-level Unicode Bidirectional Algorithm runs via `unicode-bidi`.
//!
//! The layout engine consumes bidi at run granularity: it splits each
//! paragraph into maximal same-level runs (logical order), shapes each run
//! separately (see [`crate::shape`]), then reorders runs per line at line
//! layout time. This module produces those level runs.

use unicode_bidi::{BidiInfo, Level};

/// Base paragraph direction. `Auto` derives it from the first strong
/// character per UBA rule P2/P3; `Ltr`/`Rtl` force it (Word's `w:bidi`
/// paragraph property forces RTL).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BaseDirection {
    #[default]
    Auto,
    Ltr,
    Rtl,
}

/// A maximal run of characters sharing one embedding level, in logical order.
/// `start..end` are UTF-8 byte indices into the input text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BidiRun {
    pub start: usize,
    pub end: usize,
    /// UBA embedding level; odd = right-to-left.
    pub level: u8,
}

impl BidiRun {
    pub fn is_rtl(&self) -> bool {
        level_is_rtl(self.level)
    }
}

/// One paragraph (as split by UBA rule P1 on paragraph separators) with its
/// resolved base level and level runs in logical order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BidiParagraph {
    pub start: usize,
    pub end: usize,
    /// Resolved base embedding level; odd = RTL paragraph.
    pub base_level: u8,
    pub runs: Vec<BidiRun>,
}

/// Run the UBA over `text` and return per-paragraph level runs.
///
/// DOCX paragraphs arrive one at a time in practice, but the UBA itself
/// splits on embedded paragraph separators, so the result is a vec.
pub fn bidi_paragraphs(text: &str, base: BaseDirection) -> Vec<BidiParagraph> {
    let default_level = match base {
        BaseDirection::Auto => None,
        BaseDirection::Ltr => Some(Level::ltr()),
        BaseDirection::Rtl => Some(Level::rtl()),
    };
    let info = BidiInfo::new(text, default_level);

    info.paragraphs
        .iter()
        .map(|para| {
            let mut runs: Vec<BidiRun> = Vec::new();
            for i in para.range.clone() {
                let level = info.levels[i].number();
                match runs.last_mut() {
                    Some(run) if run.level == level => run.end = i + 1,
                    _ => runs.push(BidiRun {
                        start: i,
                        end: i + 1,
                        level,
                    }),
                }
            }
            BidiParagraph {
                start: para.range.start,
                end: para.range.end,
                base_level: para.level.number(),
                runs,
            }
        })
        .collect()
}

/// True when a UBA embedding level is right-to-left.
pub fn level_is_rtl(level: u8) -> bool {
    level % 2 == 1
}

/// Visual order for a logical sequence whose elements each carry a UBA level.
///
/// The returned indices are in left-to-right paint order. This is a small
/// wrapper around `unicode-bidi`'s L2 implementation so layout code can reorder
/// shaped runs without depending directly on the transitive crate.
pub fn visual_order_for_levels(levels: &[u8]) -> Vec<usize> {
    let levels: Vec<Level> = levels
        .iter()
        .map(|&level| Level::new(level).unwrap_or_else(|_| Level::ltr()))
        .collect();
    BidiInfo::reorder_visual(&levels)
}
