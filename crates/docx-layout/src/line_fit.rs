use crate::types::{ParagraphBlock, ParagraphExtent, ParagraphFragment, TypesetRow};

pub(crate) fn trailing_leading(block: &ParagraphBlock, line: &TypesetRow) -> f64 {
    let spacing = block.attrs.as_ref().and_then(|attrs| attrs.spacing.as_ref());
    let rule = spacing.and_then(|spacing| spacing.line_rule.as_deref());
    if rule == Some("exact")
        || (rule != Some("atLeast")
            && spacing.and_then(|spacing| spacing.line_unit.as_deref()) == Some("px"))
    {
        return 0.0;
    }
    let font_height = line.ascent + line.descent;
    if !font_height.is_finite() || font_height <= 0.0 || !line.line_height.is_finite() {
        return 0.0;
    }
    let minimum = if rule == Some("atLeast") {
        spacing.and_then(|spacing| spacing.line).unwrap_or(0.0)
    } else {
        0.0
    };
    (line.line_height - font_height.max(minimum)).max(0.0)
}

pub(crate) fn slice_fit_height(block: &ParagraphBlock, lines: &[TypesetRow]) -> f64 {
    lines
        .iter()
        .map(|line| line.line_height + line.float_skip_before.unwrap_or(0.0))
        .sum::<f64>()
        - lines.last().map_or(0.0, |line| trailing_leading(block, line))
}

pub(crate) fn fragment_fit_height(
    block: &ParagraphBlock,
    measure: &ParagraphExtent,
    fragment: &ParagraphFragment,
) -> f64 {
    fragment.height
        - fragment
            .to_line
            .checked_sub(1)
            .filter(|&last| last >= fragment.from_line)
            .and_then(|last| measure.lines.get(last))
            .map_or(0.0, |line| trailing_leading(block, line))
}
