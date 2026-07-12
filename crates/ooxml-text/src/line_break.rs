use unicode_linebreak::BreakOpportunity as Uax14Opportunity;

/// A line-break opportunity at a UTF-8 boundary.
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
        .filter(|(byte_index, _)| kinsoku_allows(text, *byte_index))
        .map(|(byte_index, kind)| BreakOpportunity {
            byte_index,
            mandatory: matches!(kind, Uax14Opportunity::Mandatory),
        })
        .collect()
}

/// Apply East Asian prohibited-start and prohibited-end rules.
fn kinsoku_allows(text: &str, byte_index: usize) -> bool {
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
mod kinsoku_tests {
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
