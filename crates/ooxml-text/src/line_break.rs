//! UAX-14 line-break opportunities via the `unicode-linebreak` crate.
//!
//! Deliberate divergence from the TS `findWordBreaks` this replaces: that
//! implementation scans UTF-16 code units for spaces/hyphens, which has two
//! known gaps this module fixes by construction:
//!
//! 1. **CJK**: ideographs (UAX-14 class ID) allow a break between every pair
//!    of characters; a code-unit scan for spaces treats an unspaced CJK run
//!    as one unbreakable "word", overflowing the line. Here each
//!    inter-ideograph boundary is reported as an opportunity.
//! 2. **Surrogate safety**: indexing by UTF-16 code unit can propose a break
//!    inside a surrogate pair (emoji, supplementary-plane CJK). Opportunities
//!    here are byte indices into a `&str` and therefore always at `char`
//!    boundaries — a split surrogate is unrepresentable.
//!
//! Word-specific *no-break* refinements (e.g. `w:kinsoku` overrides,
//! non-breaking hyphens `w:noBreakHyphen`, `w:suppressAutoHyphens`) layer on
//! top of these raw UAX-14 opportunities — see [`crate::word_metrics`].

use unicode_linebreak::BreakOpportunity as Uax14Opportunity;

/// One line-break opportunity.
///
/// `byte_index` is the UTF-8 byte position where the next line would start
/// (i.e. the break is *before* this index's character), matching
/// `unicode-linebreak` semantics. Always a `char` boundary. Per UAX-14 the
/// end of text is reported as a final mandatory break at `text.len()`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BreakOpportunity {
    pub byte_index: usize,
    /// `true` for mandatory breaks (LF, paragraph separator, end of text),
    /// `false` for allowed (optional) break opportunities.
    pub mandatory: bool,
}

/// All UAX-14 break opportunities for `text`, in ascending byte order.
pub fn break_opportunities(text: &str) -> Vec<BreakOpportunity> {
    unicode_linebreak::linebreaks(text)
        .filter(|(byte_index, _)| word_kinsoku_allows(text, *byte_index))
        .map(|(byte_index, kind)| BreakOpportunity {
            byte_index,
            mandatory: matches!(kind, Uax14Opportunity::Mandatory),
        })
        .collect()
}

/// Word's default East Asian prohibited-start/prohibited-end refinement.
/// UAX-14 supplies the broad opportunity set; kinsoku removes breaks that
/// would strand opening punctuation at line end or closing punctuation at
/// line start. The terminal break is always retained.
fn word_kinsoku_allows(text: &str, byte_index: usize) -> bool {
    if byte_index == 0 || byte_index >= text.len() {
        return true;
    }
    const PROHIBITED_START: &[char] = &[
        '!', '%', ')', ',', '.', ':', ';', '?', ']', '}', '¢', '°', '’', '”', '†', '‡', '…', '‰',
        '′', '″', '℃', '、', '。', '〉', '》', '」', '』', '】', '〕', '〗', '〙', '〛', '゛',
        '゜', 'ゝ', 'ゞ', 'ー', 'ァ', 'ィ', 'ゥ', 'ェ', 'ォ', 'ッ', 'ャ', 'ュ', 'ョ', 'ヮ', 'ヵ',
        'ヶ', '・', 'ヽ', 'ヾ', '！', '％', '）', '，', '．', '：', '；', '？', '］', '｝',
    ];
    const PROHIBITED_END: &[char] = &[
        '(', '[', '{', '£', '¥', '‘', '“', '〈', '《', '「', '『', '【', '〔', '〖', '〘', '〚',
        '（', '［', '｛', '￥',
    ];
    let previous = text[..byte_index].chars().next_back();
    let next = text[byte_index..].chars().next();
    !previous.is_some_and(|ch| PROHIBITED_END.contains(&ch))
        && !next.is_some_and(|ch| PROHIBITED_START.contains(&ch))
}

#[cfg(test)]
mod word_tests {
    use super::*;

    #[test]
    fn kinsoku_filters_breaks_before_closing_and_after_opening_punctuation() {
        let text = "漢（字）漢";
        let breaks = break_opportunities(text);
        let indices: Vec<usize> = breaks.iter().map(|item| item.byte_index).collect();
        let after_open = "漢（".len();
        let before_close = "漢（字".len();
        assert!(!indices.contains(&after_open));
        assert!(!indices.contains(&before_close));
        assert_eq!(indices.last().copied(), Some(text.len()));
    }
}
