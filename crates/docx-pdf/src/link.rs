//! PDF link annotations.
//!
//! [`collect_links`] turns every visible primitive that carries an external
//! hyperlink into a clickable rectangle in PDF points, and [`write_annotation`]
//! serializes one as a `/Link` annotation with a `/URI` action. Internal
//! anchors are skipped: this writer has no named destination table.

use docx_layout::display_list::{
    DisplayPage, DocAttrs, GlyphRunPrimitive, Primitive, TextRunPrimitive,
};
use pdf_writer::types::{ActionType, AnnotationType};
use pdf_writer::{Name, Pdf, Rect, Ref, Str, TextStr};

use crate::units::{PX_TO_PT, font_size_px};

/// One clickable rectangle on a PDF page.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct LinkAnnotation {
    /// PDF points, `[x0, y0, x1, y1]` with a bottom-left origin.
    pub rect: [f32; 4],
    /// What a click on this rectangle does.
    pub action: LinkAction,
    /// Optional annotation title written as `/T`.
    pub title: Option<String>,
}

/// The action a link annotation performs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum LinkAction {
    /// Open an external URI.
    Uri(String),
}

/// Collect the clickable regions of one page.
///
/// Body, header, footer, note-area and note-separator primitives are all
/// scanned. Rectangles are converted from the display list's top-left CSS
/// pixel space to PDF points with a bottom-left origin.
pub(crate) fn collect_links(page: &DisplayPage) -> Vec<LinkAnnotation> {
    let Some(page_height) = page.height.as_f64() else {
        return Vec::new();
    };

    let mut links = Vec::new();
    let regions = page
        .header
        .iter()
        .map(|region| region.primitives.as_slice())
        .chain(
            page.footer
                .iter()
                .map(|region| region.primitives.as_slice()),
        )
        .chain(
            page.note_areas
                .iter()
                .map(|region| region.primitives.as_slice()),
        )
        .chain(
            page.note_areas
                .iter()
                .map(|region| region.separator_primitives.as_slice()),
        );
    for primitives in std::iter::once(page.primitives.as_slice()).chain(regions) {
        for primitive in primitives {
            let Some((attrs, rect_px)) = primitive_rect(primitive) else {
                continue;
            };
            if attrs.hidden_object == Some(true) {
                continue;
            }
            let Some(action) = link_action(attrs) else {
                continue;
            };
            let Some(rect) = to_pdf_rect(rect_px, page_height) else {
                continue;
            };
            links.push(LinkAnnotation {
                rect,
                action,
                title: link_title(attrs),
            });
        }
    }
    merge_adjacent_links(links)
}

/// Merge annotations that share an action and title and whose rectangles
/// touch or overlap. A hyperlink split across shaped primitive fragments must
/// become one clickable box, not one annotation per fragment.
fn merge_adjacent_links(links: Vec<LinkAnnotation>) -> Vec<LinkAnnotation> {
    let mut merged: Vec<LinkAnnotation> = Vec::with_capacity(links.len());
    for link in links {
        if let Some(existing) = merged.iter_mut().find(|existing| {
            existing.action == link.action
                && existing.title == link.title
                && rects_touch(existing.rect, link.rect)
        }) {
            for (index, value) in existing.rect.iter_mut().enumerate() {
                *value = if index < 2 {
                    value.min(link.rect[index])
                } else {
                    value.max(link.rect[index])
                };
            }
        } else {
            merged.push(link);
        }
    }
    merged
}

/// `true` when two PDF rectangles overlap or are within two points.
fn rects_touch(left: [f32; 4], right: [f32; 4]) -> bool {
    const SLOP: f32 = 2.0;
    left[0] <= right[2] + SLOP
        && right[0] <= left[2] + SLOP
        && left[1] <= right[3] + SLOP
        && right[1] <= left[3] + SLOP
}

/// Write one `/Link` annotation object for `page`.
pub(crate) fn write_annotation(pdf: &mut Pdf, id: Ref, page: Ref, annotation: &LinkAnnotation) {
    let mut writer = pdf.annotation(id);
    writer
        .subtype(AnnotationType::Link)
        .rect(Rect::new(
            annotation.rect[0],
            annotation.rect[1],
            annotation.rect[2],
            annotation.rect[3],
        ))
        .page(page);
    if let Some(title) = annotation.title.as_deref() {
        writer.pair(Name(b"T"), TextStr(title));
    }
    let mut action = writer.action();
    match &annotation.action {
        LinkAction::Uri(uri) => {
            action.action_type(ActionType::Uri);
            action.uri(Str(uri.as_bytes()));
        }
    }
}

/// The link metadata of one primitive and its clickable box in CSS pixels,
/// `[left, top, right, bottom]`.
fn primitive_rect(primitive: &Primitive) -> Option<(&DocAttrs, [f64; 4])> {
    match primitive {
        Primitive::Text(text) => {
            if text.hidden {
                return None;
            }
            Some((&text.attrs, text_rect(text)?))
        }
        Primitive::GlyphRun(run) => {
            if run.hidden {
                return None;
            }
            Some((&run.attrs, glyph_run_rect(run)?))
        }
        Primitive::Image(image) => Some((
            &image.attrs,
            frame(
                image.x.as_f64()?,
                image.y.as_f64()?,
                image.w.as_f64()?,
                image.h.as_f64()?,
            )?,
        )),
        Primitive::Shape(shape) => Some((
            &shape.attrs,
            frame(
                shape.x.as_f64()?,
                shape.y.as_f64()?,
                shape.w.as_f64()?,
                shape.h.as_f64()?,
            )?,
        )),
        Primitive::Rect(rect) => Some((
            &rect.attrs,
            frame(
                rect.x.as_f64()?,
                rect.y.as_f64()?,
                rect.w.as_f64()?,
                rect.h.as_f64()?,
            )?,
        )),
        Primitive::Decoration(decoration) => Some((
            &decoration.attrs,
            frame(
                decoration.x.as_f64()?,
                decoration.y.as_f64()?,
                decoration.w.as_f64()?,
                decoration.h.as_f64()?,
            )?,
        )),
        // A line's box is degenerate in one axis, so it never yields a
        // clickable area; hyperlinked rules are not emitted by DOCX anyway.
        Primitive::Line(_) => None,
    }
}

/// Box of a fallback text run around its baseline, in CSS pixels.
fn text_rect(text: &TextRunPrimitive) -> Option<[f64; 4]> {
    let x = text.x.as_f64()?;
    let baseline = text.baseline_y.as_f64()?;
    let width = text.width.as_f64()?;
    let size = f64::from(font_size_px(&text.font));
    Some([x, baseline - size * 0.8, x + width, baseline + size * 0.25])
}

/// Union of the glyph boxes of one shaped run, in CSS pixels.
fn glyph_run_rect(run: &GlyphRunPrimitive) -> Option<[f64; 4]> {
    let size = run.size;
    if !size.is_finite() || size <= 0.0 {
        return None;
    }
    let mut bounds: Option<[f64; 4]> = None;
    for glyph in &run.glyphs {
        let (x, y, advance) = (glyph.x, glyph.y, glyph.advance);
        if !x.is_finite() || !y.is_finite() || !advance.is_finite() {
            continue;
        }
        let left = x.min(x + advance);
        let right = x.max(x + advance);
        bounds = Some(match bounds {
            None => [left, y, right, y],
            Some(bounds) => [
                bounds[0].min(left),
                bounds[1].min(y),
                bounds[2].max(right),
                bounds[3].max(y),
            ],
        });
    }
    let bounds = bounds?;
    Some([
        bounds[0],
        bounds[1] - size * 0.8,
        bounds[2],
        bounds[3] + size * 0.25,
    ])
}

/// `[x, y, x + w, y + h]` in CSS pixels.
fn frame(x: f64, y: f64, w: f64, h: f64) -> Option<[f64; 4]> {
    let rect = [x, y, x + w, y + h];
    rect.iter().all(|value| value.is_finite()).then_some(rect)
}

/// Resolve the external hyperlink destination of a primitive.
///
/// Internal anchors (`#name`) and `link_doc_location` have no named
/// destination table in this writer, so they are skipped rather than emitted
/// as `/GoTo` actions that cannot resolve.
fn link_action(attrs: &DocAttrs) -> Option<LinkAction> {
    let href = attrs
        .href
        .as_deref()
        .map(str::trim)
        .filter(|href| !href.is_empty())?;
    if href.starts_with('#') {
        return None;
    }
    Some(LinkAction::Uri(href.to_string()))
}

/// Tooltip/`w:tooltip` of the hyperlink, if any.
fn link_title(attrs: &DocAttrs) -> Option<String> {
    attrs
        .link_title
        .as_deref()
        .or(attrs.tooltip.as_deref())
        .map(str::trim)
        .filter(|title| !title.is_empty())
        .map(str::to_string)
}

/// Convert a `[left, top, right, bottom]` CSS-pixel box to PDF points with a
/// bottom-left origin, rejecting degenerate boxes.
fn to_pdf_rect(rect_px: [f64; 4], page_height: f64) -> Option<[f32; 4]> {
    if !rect_px.iter().all(|value| value.is_finite()) || !page_height.is_finite() {
        return None;
    }
    let left = rect_px[0].min(rect_px[2]);
    let right = rect_px[0].max(rect_px[2]);
    let top = rect_px[1].min(rect_px[3]);
    let bottom = rect_px[1].max(rect_px[3]);
    let scale = f64::from(PX_TO_PT);
    let rect = [
        (left * scale) as f32,
        ((page_height - bottom) * scale) as f32,
        (right * scale) as f32,
        ((page_height - top) * scale) as f32,
    ];
    if !rect.iter().all(|value| value.is_finite()) || rect[2] <= rect[0] || rect[3] <= rect[1] {
        return None;
    }
    Some(rect)
}
