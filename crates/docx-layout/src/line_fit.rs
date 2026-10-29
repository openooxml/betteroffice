use crate::types::{
    ParagraphAttrs, ParagraphBlock, ParagraphExtent, ParagraphFragment, Run, RunFormatting,
    TypesetRow, UnderlineSpec,
};

/// Eligible: measured text, fields, breaks, font/color/bidi/spacing metadata,
/// font-box highlights, bounded single underlines/strikes and geometric tab leaders.
/// Other painted features require full-height fitting.
fn eligible_painted_descent(block: &ParagraphBlock, line: &TypesetRow) -> Option<f64> {
    let mut tabs = None;
    if let Some(attrs) = &block.attrs {
        let ParagraphAttrs {
            justification_mode: _,
            has_source_content: _,
            horizontal_rules,
            alignment: _,
            spacing: _,
            spacing_explicit: _,
            indent: _,
            keep_next: _,
            keep_lines: _,
            widow_control: _,
            page_break_before: _,
            page_break_before_run: _,
            style_id: _,
            effective_style_id: _,
            contextual_spacing: _,
            bidi: _,
            borders,
            shading: None,
            tabs: paragraph_tabs,
            num_pr: _,
            list_marker,
            list_is_bullet: _,
            list_marker_hidden,
            list_marker_font_family: _,
            list_marker_font_size: _,
            list_marker_bold: _,
            list_marker_italic: _,
            list_marker_color: _,
            list_marker_suffix: _,
            list_marker_revision: _,
            default_tab_stop_twips: _,
            default_font_size: _,
            default_font_family: _,
            suppress_empty_paragraph_height: _,
            snap_to_grid: _,
            auto_space_de: _,
            auto_space_dn: _,
            doc_grid_pitch_px: _,
            p_pr_ins: None,
            p_pr_del: None,
        } = attrs
        else {
            return None;
        };
        if !horizontal_rules.is_empty()
            || (list_marker
                .as_ref()
                .is_some_and(|marker| !marker.is_empty())
                && *list_marker_hidden != Some(true))
        {
            return None;
        }
        if let Some(borders) = borders
            && [
                borders.top.as_ref(),
                borders.bottom.as_ref(),
                borders.between.as_ref(),
                borders.left.as_ref(),
                borders.right.as_ref(),
                borders.bar.as_ref(),
            ]
            .into_iter()
            .flatten()
            .any(|border| {
                !matches!(border.style.as_deref(), Some("none" | "nil"))
                    && border.width.unwrap_or(1.0) > 0.0
            })
        {
            return None;
        }
        tabs = paragraph_tabs.as_ref();
    }

    let mut bottom = line.descent;
    for run in block.runs.get(line.head_run..=line.tail_run)? {
        let (fmt, tab) = match run {
            Run::Text(run) if run.inline_sdt_widget.is_none() => (&run.fmt, None),
            Run::Field(run) => (&run.fmt, None),
            Run::Tab(run) if run.leader_glyphs.is_none() => (&run.fmt, Some(run)),
            Run::LineBreak(_) => continue,
            _ => return None,
        };
        let RunFormatting {
            bold: _,
            italic: _,
            underline,
            strike,
            color: _,
            highlight,
            font_family: _,
            font_slots: _,
            font_size,
            font_size_cs,
            bold_cs: _,
            italic_cs: _,
            complex_script: _,
            language: _,
            letter_spacing: _,
            superscript: None | Some(false),
            subscript: None | Some(false),
            all_caps: _,
            small_caps: _,
            position_px: None,
            horizontal_scale: _,
            kerning_min_pt: _,
            imprint: None | Some(false),
            emboss: None | Some(false),
            text_shadow: None | Some(false),
            text_outline: None | Some(false),
            emphasis_mark: None,
            hidden,
            rtl: _,
            snap_to_grid: _,
            text_effect: None,
            modern_effects: None,
            hyperlink: _,
            footnote_ref_id: _,
            endnote_ref_id: _,
            comment_ids,
            is_insertion: None | Some(false),
            is_deletion: None | Some(false),
            change_author: _,
            change_date: _,
            change_revision_id: _,
            logical_order: _,
            bidi_level: _,
        } = fmt
        else {
            return None;
        };
        if comment_ids.as_ref().is_some_and(|ids| !ids.is_empty()) {
            return None;
        }
        let painted_size = font_size.unwrap_or(11.0);
        if font_size_cs.is_some_and(|size| size != painted_size) {
            return None;
        }
        let font_px = painted_size * 96.0 / 72.0;
        if !font_px.is_finite() || font_px <= 0.0 {
            return None;
        }
        let thickness = (font_px / 16.0).round().max(1.0);
        let underline_bottom = (font_px * 0.1875).round() + thickness;
        let underlined = match underline {
            None | Some(UnderlineSpec::Flag(false)) => false,
            Some(UnderlineSpec::Flag(true)) => true,
            Some(UnderlineSpec::Styled { style, color: _ }) => match style.as_deref() {
                None | Some("single" | "solid") => true,
                _ => return None,
            },
        };
        if underlined || *hidden == Some(true) {
            bottom = bottom.max(underline_bottom);
        }
        if *strike == Some(true) {
            bottom = bottom.max(-(font_px * 0.3).round() + thickness);
        }
        if highlight.is_some() {
            bottom = bottom.max(font_px * 0.2);
        }
        if tab.is_some() {
            for stop in tabs.into_iter().flatten() {
                if stop.val == "clear" {
                    continue;
                }
                let leader_bottom = match stop.leader.as_deref() {
                    None | Some("none") => continue,
                    Some("dot" | "middleDot" | "hyphen") => -(font_px * 0.25).round() + thickness,
                    Some("underscore") => underline_bottom,
                    Some("heavy") => (font_px * 0.1875).round() + thickness.max(2.0),
                    _ => return None,
                };
                bottom = bottom.max(leader_bottom);
            }
        }
    }
    Some(bottom)
}

pub(crate) fn trailing_leading(block: &ParagraphBlock, line: &TypesetRow) -> f64 {
    let spacing = block
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.spacing.as_ref());
    let rule = spacing.and_then(|spacing| spacing.line_rule.as_deref());
    let unit = spacing.and_then(|spacing| spacing.line_unit.as_deref());
    if !matches!(rule, None | Some("auto" | "atLeast"))
        || !matches!(unit, None | Some("multiplier" | "px"))
        || (rule != Some("atLeast") && unit == Some("px"))
    {
        return 0.0;
    }
    let font_height = line.ascent + line.descent;
    if !font_height.is_finite()
        || font_height <= 0.0
        || line.ascent < 0.0
        || line.descent < 0.0
        || !line.line_height.is_finite()
    {
        return 0.0;
    }
    let Some(painted_descent) = eligible_painted_descent(block, line) else {
        return 0.0;
    };
    let minimum = if rule == Some("atLeast") {
        spacing.and_then(|spacing| spacing.line).unwrap_or(0.0)
    } else {
        0.0
    };
    if !minimum.is_finite() || minimum < 0.0 {
        return 0.0;
    }
    (line.line_height - (line.ascent + line.descent.max(painted_descent)).max(minimum)).max(0.0)
}

pub(crate) fn slice_fit_height(block: &ParagraphBlock, lines: &[TypesetRow]) -> f64 {
    lines
        .iter()
        .map(|line| line.line_height + line.float_skip_before.unwrap_or(0.0))
        .sum::<f64>()
        - lines
            .last()
            .map_or(0.0, |line| trailing_leading(block, line))
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
