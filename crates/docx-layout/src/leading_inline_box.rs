use std::borrow::Cow;

use crate::types::{
    LeadingInlineBox, ParagraphBlock, ParagraphExtent, ReservationRef, ReservedRun, Run,
};

fn size(paragraph: &ParagraphBlock) -> Option<&LeadingInlineBox> {
    paragraph
        .attrs
        .as_ref()?
        .leading_inline_box
        .as_ref()
        .filter(|size| {
            size.width.is_finite()
                && size.width > 0.0
                && size.height.is_finite()
                && size.height > 0.0
                && (size.width as f32).is_finite()
                && (size.height as f32).is_finite()
        })
}

pub(crate) fn measurement_paragraph(paragraph: &ParagraphBlock) -> Cow<'_, ParagraphBlock> {
    let Some(size) = size(paragraph) else {
        return Cow::Borrowed(paragraph);
    };
    let mut measured = paragraph.clone();
    measured.attrs.as_mut().unwrap().leading_inline_box = None;
    measured.runs.insert(
        0,
        Run::Reserved(ReservedRun {
            reservation: ReservationRef {
                owner: "numbering".to_owned(),
                id: "firstPictureBullet".to_owned(),
            },
            width: size.width,
            height: size.height,
            pm_start: None,
            pm_end: None,
        }),
    );
    Cow::Owned(measured)
}

fn source_range(
    head_run: &mut usize,
    head_char: &mut usize,
    tail_run: &mut usize,
    tail_char: &mut usize,
    run_count: usize,
) {
    if *tail_run == 0 {
        *head_run = run_count;
        *tail_run = run_count;
        *head_char = 0;
        *tail_char = 0;
    } else {
        if *head_run == 0 {
            *head_char = 0;
        }
        *head_run = head_run.saturating_sub(1);
        *tail_run -= 1;
    }
}

pub(crate) fn restore_extent(paragraph: &ParagraphBlock, extent: &mut ParagraphExtent) {
    let Some(size) = size(paragraph) else {
        return;
    };
    let is_rtl = paragraph.attrs.as_ref().and_then(|attrs| attrs.bidi) == Some(true);
    for line in &mut extent.lines {
        let has_box = line.head_run == 0;
        if has_box {
            line.width = (line.width - size.width).max(0.0);
            let offset = if is_rtl {
                &mut line.right_offset
            } else {
                &mut line.left_offset
            };
            *offset = Some(offset.unwrap_or(0.0) + size.width);
            let height = line.line_height;
            line.ascent = line.ascent.max(size.height);
            line.line_height = line.line_height.max(line.ascent + line.descent);
            extent.total_height += line.line_height - height;
        }
        source_range(
            &mut line.head_run,
            &mut line.head_char,
            &mut line.tail_run,
            &mut line.tail_char,
            paragraph.runs.len(),
        );
        for segment in line.segments.iter_mut().flatten() {
            if segment.head_run == 0 {
                if !is_rtl {
                    segment.left_offset += size.width;
                }
                segment.available_width = (segment.available_width - size.width).max(0.0);
                segment.width = (segment.width - size.width).max(0.0);
            }
            source_range(
                &mut segment.head_run,
                &mut segment.head_char,
                &mut segment.tail_run,
                &mut segment.tail_char,
                paragraph.runs.len(),
            );
        }
        if let Some(advances) = &mut line.run_advances {
            advances.retain_mut(|advance| {
                advance.run_index = advance.run_index.and_then(|index| index.checked_sub(1));
                advance.run_index.is_some()
            });
        }
        if let Some(clusters) = &mut line.cluster_advances {
            clusters.retain_mut(|cluster| {
                cluster.run_index = cluster.run_index.and_then(|index| index.checked_sub(1));
                if has_box && !is_rtl {
                    cluster.x_offset = cluster
                        .x_offset
                        .map(|offset| (offset - size.width).max(0.0));
                }
                cluster.run_index.is_some()
            });
        }
        if let Some(slices) = &mut line.bidi_slices {
            slices.retain_mut(|slice| {
                slice.run_index = slice.run_index.and_then(|index| index.checked_sub(1));
                slice.run_index.is_some()
            });
        }
        if let Some(packed) = line.glyph_runs.as_deref() {
            let shift = (has_box && !is_rtl).then_some(size.width as f32);
            line.glyph_runs = without_leading_box(packed, shift);
        }
    }
}

/// The packed glyph runs of a line without the leading box's reserved run:
/// the others move down one run index, and `shift` moves them left on the line
/// when the box sat before them. `None` when no run is left.
fn without_leading_box(packed: &str, shift: Option<f32>) -> Option<String> {
    let mut runs = ooxml_text::unpack_glyph_runs(packed)?;
    runs.retain_mut(|run| {
        let Some(index) = run.head.run_index.checked_sub(1) else {
            return false;
        };
        run.head.run_index = index;
        if let Some(shift) = shift {
            run.head.x_offset = (run.head.x_offset - shift).max(0.0);
        }
        true
    });
    ooxml_text::pack_glyph_runs(&runs)
}

#[cfg(test)]
mod glyph_run_tests {
    use super::without_leading_box;
    use ooxml_text::{
        GlyphRunHead, MeasuredGlyph, MeasuredGlyphRun, pack_glyph_runs, unpack_glyph_runs,
    };

    fn run(run_index: u32, x_offset: f32, start: u32, bidi_level: u8) -> MeasuredGlyphRun {
        let rtl = bidi_level % 2 == 1;
        let glyphs = (0..4u32)
            .map(|index| MeasuredGlyph {
                id: 40 + index,
                cluster: start + if rtl { 3 - index } else { index },
                advance: 5.5,
                dx: if rtl && index == 1 { -1.25 } else { 0.0 },
                dy: if rtl && index == 1 { 2.5 } else { 0.0 },
            })
            .collect();
        MeasuredGlyphRun {
            head: GlyphRunHead {
                run_index,
                start_char: start,
                end_char: start + 4,
                font_id: 3,
                bidi_level,
                x_offset,
                advance: 22.0,
            },
            glyphs,
        }
    }

    fn expected(runs: &[MeasuredGlyphRun], shift: f32) -> Vec<MeasuredGlyphRun> {
        runs.iter()
            .map(|run| {
                let mut run = run.clone();
                run.head.run_index -= 1;
                run.head.x_offset = (run.head.x_offset - shift).max(0.0);
                run
            })
            .collect()
    }

    #[test]
    fn a_shifted_line_keeps_its_latin_and_arabic_glyphs() {
        for level in [0, 1] {
            let runs = [run(1, 30.0, 0, level), run(2, 52.0, 4, level)];
            let packed = pack_glyph_runs(&runs).unwrap();
            let shifted = without_leading_box(&packed, Some(30.0)).unwrap();
            assert_eq!(unpack_glyph_runs(&shifted).unwrap(), expected(&runs, 30.0));
        }
    }

    #[test]
    fn without_a_shift_only_the_run_indexes_move() {
        let runs = [run(1, 12.0, 0, 1), run(2, 34.0, 4, 0)];
        let packed = pack_glyph_runs(&runs).unwrap();
        let moved = without_leading_box(&packed, None).unwrap();
        assert_eq!(unpack_glyph_runs(&moved).unwrap(), expected(&runs, 0.0));
    }

    #[test]
    fn the_reserved_run_is_dropped_and_an_empty_line_packs_to_nothing() {
        let reserved = run(0, 0.0, 0, 0);
        let kept = run(1, 30.0, 0, 0);
        let packed = pack_glyph_runs(&[reserved.clone(), kept.clone()]).unwrap();
        let moved = without_leading_box(&packed, None).unwrap();
        assert_eq!(
            unpack_glyph_runs(&moved).unwrap(),
            expected(std::slice::from_ref(&kept), 0.0)
        );
        let alone = pack_glyph_runs(&[reserved]).unwrap();
        assert_eq!(without_leading_box(&alone, None), None);
    }
}
