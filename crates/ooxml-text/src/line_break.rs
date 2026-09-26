//! UAX-14 line breaks with kinsoku filtering and Unicode scalar boundaries.

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

/// UAX-14 keeps a hyphen glued to the number after it, so `COVID-19` never
/// wraps. PowerPoint breaks there, as its own render of the corpus shows, so a
/// slide takes this set instead: the UAX-14 opportunities plus one after every
/// hyphen that a digit follows.
pub fn presentation_break_opportunities(text: &str) -> Vec<BreakOpportunity> {
    let mut opportunities = break_opportunities(text);
    let mut extra: Vec<usize> = Vec::new();
    for (byte_index, character) in text.char_indices() {
        if !matches!(character, '-' | '\u{2010}') {
            continue;
        }
        let after = byte_index + character.len_utf8();
        let follows_word = text[..byte_index]
            .chars()
            .next_back()
            .is_some_and(char::is_alphanumeric);
        let leads_digit = text[after..]
            .chars()
            .next()
            .is_some_and(|next| next.is_ascii_digit());
        if follows_word && leads_digit {
            extra.push(after);
        }
    }
    if extra.is_empty() {
        return opportunities;
    }
    extra.retain(|index| !opportunities.iter().any(|value| value.byte_index == *index));
    opportunities.extend(extra.into_iter().map(|byte_index| BreakOpportunity {
        byte_index,
        mandatory: false,
    }));
    opportunities.sort_by_key(|value| value.byte_index);
    opportunities
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
    fn a_slide_breaks_after_a_hyphen_that_a_digit_follows() {
        let indexes = |text: &str| {
            presentation_break_opportunities(text)
                .into_iter()
                .map(|value| value.byte_index)
                .collect::<Vec<_>>()
        };
        assert_eq!(indexes("COVID-19 on"), [6, 9, 11]);
        assert_eq!(indexes("2020-2021 x"), [5, 10, 11]);
        assert_eq!(indexes("well-known thing"), [5, 11, 16]);
        assert_eq!(indexes("-19 x"), [4, 5]);
    }

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
