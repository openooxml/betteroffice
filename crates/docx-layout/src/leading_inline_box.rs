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
    }
}
