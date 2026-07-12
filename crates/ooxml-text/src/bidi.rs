//! Unicode Bidirectional Algorithm runs.

use unicode_bidi::{BidiInfo, Level};

/// Base paragraph direction.
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

/// A paragraph's resolved bidi runs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BidiParagraph {
    pub start: usize,
    pub end: usize,
    /// Resolved base embedding level; odd = RTL paragraph.
    pub base_level: u8,
    pub runs: Vec<BidiRun>,
}

/// Resolve bidi runs for `text`.
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

/// Return visual indices for logical embedding levels.
pub fn visual_order_for_levels(levels: &[u8]) -> Vec<usize> {
    let levels: Vec<Level> = levels
        .iter()
        .map(|&level| Level::new(level).unwrap_or_else(|_| Level::ltr()))
        .collect();
    BidiInfo::reorder_visual(&levels)
}
