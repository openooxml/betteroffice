//! PDF link annotations.
//!
//! [`collect_links`] turns every visible primitive that carries an external
//! hyperlink into a clickable rectangle in PDF points, and [`write_annotation`]
//! serializes one as a `/Link` annotation with a `/URI` action. Internal
//! anchors are skipped: this writer has no named destination table.
//!
//! PDF/A-2u additionally requires the annotation's `/F` Print flag
//! (ISO 19005-2, 6.3.2) but exempts `/Link` annotations from the `/AP`
//! appearance-stream requirement (6.3.3), so archival output sets `/F 4` and
//! writes no appearance dictionary.

use docx_layout::display_list::{
    DisplayPage, DocAttrs, GlyphRunPrimitive, ImagePrimitive, Primitive, TextRunPrimitive,
};
use pdf_writer::types::{ActionType, AnnotationFlags, AnnotationType};
use pdf_writer::{Name, Pdf, Rect, Ref, Str, TextStr};

use crate::pdfa_limits::PDFA_STRING_TOO_LONG;
use crate::pdfa_limits::{bytes_too_long, text_too_long};

use crate::units::{PX_TO_PT, font_size_px};
use crate::{PdfConformance, PdfError};

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
    /// An action this writer cannot serialize, named by its PDF action type
    /// (for example `"JavaScript"`). [`collect_links`] emits only URIs; the
    /// variant exists so [`write_annotation`] can reject such input instead of
    /// emitting a link with no action or silently dropping the click target.
    #[cfg_attr(not(test), allow(dead_code))]
    Unsupported(&'static str),
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
    // A fragment that bridges two already-collected boxes can appear after
    // both of them. Re-check until the list is stable so one hyperlink never
    // leaves multiple adjacent annotations.
    loop {
        let mut changed = false;
        'pairs: for left in 0..merged.len() {
            for right in left + 1..merged.len() {
                if merged[left].action == merged[right].action
                    && merged[left].title == merged[right].title
                    && rects_touch(merged[left].rect, merged[right].rect)
                {
                    let other = merged.remove(right);
                    let existing = &mut merged[left];
                    for (index, value) in existing.rect.iter_mut().enumerate() {
                        *value = if index < 2 {
                            value.min(other.rect[index])
                        } else {
                            value.max(other.rect[index])
                        };
                    }
                    changed = true;
                    break 'pairs;
                }
            }
        }
        if !changed {
            break;
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
///
/// The rectangle, title and URI action are written exactly as collected.
/// [`PdfConformance::PdfA2u`] adds `/F 4`: ISO 19005-2, 6.3.2 requires the
/// Print flag and forbids Hidden, Invisible, NoView and ToggleNoView, which a
/// zero `Print`-only value satisfies. Link annotations are exempt from the
/// appearance-stream requirement (6.3.3), so no `/AP` dictionary is written.
/// `Standard` keeps the historical dictionary without `/F`.
///
/// An action this writer cannot serialize is rejected with
/// [`PdfError::Unsupported`]; it is never dropped or downgraded to a dead
/// click target.
pub(crate) fn write_annotation(
    pdf: &mut Pdf,
    id: Ref,
    page: Ref,
    annotation: &LinkAnnotation,
    conformance: PdfConformance,
) -> Result<(), PdfError> {
    let uri = match &annotation.action {
        LinkAction::Uri(uri) => uri,
        LinkAction::Unsupported(kind) => return Err(PdfError::Unsupported(*kind)),
    };
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
    if conformance.is_pdfa() {
        writer.flags(AnnotationFlags::PRINT);
        if let Some(title) = annotation.title.as_deref() {
            if text_too_long(title) {
                return Err(crate::PdfError::Unsupported(PDFA_STRING_TOO_LONG));
            }
        }
        if bytes_too_long(uri.as_bytes()) {
            return Err(crate::PdfError::Unsupported(PDFA_STRING_TOO_LONG));
        }
    }
    if let Some(title) = annotation.title.as_deref() {
        writer.pair(Name(b"T"), TextStr(title));
    }
    let mut action = writer.action();
    action.action_type(ActionType::Uri);
    action.uri(Str(uri.as_bytes()));
    Ok(())
}

/// The link metadata of one primitive and its clickable box in CSS pixels,
/// `[left, top, right, bottom]`.
fn primitive_rect(primitive: &Primitive) -> Option<(&DocAttrs, [f64; 4])> {
    match primitive {
        Primitive::Text(text) => {
            if text.hidden {
                return None;
            }
            Some((&text.attrs, text_link_rect(text)?))
        }
        Primitive::GlyphRun(run) => {
            if run.hidden {
                return None;
            }
            Some((&run.attrs, glyph_run_rect(run)?))
        }
        Primitive::Image(image) => Some((
            &image.attrs,
            rotated_box(image_frame(image)?, image.rotation_deg.as_ref()),
        )),
        Primitive::Shape(shape) => {
            let rotation = shape
                .transform
                .as_ref()
                .and_then(|transform| transform.rotation.as_ref());
            Some((
                &shape.attrs,
                rotated_box(
                    frame(
                        shape.x.as_f64()?,
                        shape.y.as_f64()?,
                        shape.w.as_f64()?,
                        shape.h.as_f64()?,
                    )?,
                    rotation,
                ),
            ))
        }
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

/// Bounds of a text run after the painter's rotation.
///
/// The painter rotates the run about `(x + width / 2, baseline - 0.2em)`,
/// which is the baseline offset it uses for every text matrix, rather than
/// about the visual box centre.
fn text_link_rect(text: &TextRunPrimitive) -> Option<[f64; 4]> {
    let rect = text_rect(text)?;
    let Some(degrees) = rotation_degrees(text.rotation_deg.as_ref()) else {
        return Some(rect);
    };
    let baseline = text.baseline_y.as_f64()?;
    let size = f64::from(font_size_px(&text.font));
    Some(rotate_box(
        rect,
        degrees,
        ((rect[0] + rect[2]) / 2.0, baseline - size * 0.2),
    ))
}

/// Bounds of one shaped run after the painter's placement transform.
///
/// Both paths scale each glyph box about the run's left edge first, then a
/// rotated run turns the scaled points about the glyph-rect centre
/// `(mid-x, first baseline - 0.2em)`. That is the transform the raster and
/// PDF painters apply to every glyph outline.
fn glyph_run_rect(run: &GlyphRunPrimitive) -> Option<[f64; 4]> {
    let size = run.size;
    if !size.is_finite() || size <= 0.0 {
        return None;
    }
    let scale = match run.horizontal_scale.as_ref() {
        Some(value) => {
            let value = value.as_f64()?;
            if !value.is_finite() || value <= 0.0 {
                return None;
            }
            value / 100.0
        }
        None => 1.0,
    };
    let first = run.glyphs.first()?;
    if !first.x.is_finite() || !first.y.is_finite() {
        return None;
    }
    let rotation = run
        .rotation_deg
        .as_ref()
        .and_then(serde_json::Number::as_f64)
        .unwrap_or(0.0);
    let rotated = rotation.is_finite() && rotation.abs() > f64::EPSILON;
    let (sin, cos) = rotation.to_radians().sin_cos();

    let mut left = f64::INFINITY;
    let mut right = f64::NEG_INFINITY;
    for glyph in &run.glyphs {
        if !glyph.x.is_finite() || !glyph.advance.is_finite() {
            continue;
        }
        left = left.min(glyph.x);
        right = right.max(glyph.x + glyph.advance);
    }
    if !left.is_finite() || !right.is_finite() {
        return None;
    }
    let pivot_x = (left + right) / 2.0;
    let pivot_y = first.y - size * 0.2;

    let mut bounds = [
        f64::INFINITY,
        f64::INFINITY,
        f64::NEG_INFINITY,
        f64::NEG_INFINITY,
    ];
    for glyph in &run.glyphs {
        if !glyph.x.is_finite() || !glyph.y.is_finite() || !glyph.advance.is_finite() {
            continue;
        }
        let top = glyph.y - size * 0.8;
        let bottom = glyph.y + size * 0.25;
        for (x, y) in [
            (glyph.x, top),
            (glyph.x + glyph.advance, top),
            (glyph.x, bottom),
            (glyph.x + glyph.advance, bottom),
        ] {
            let (x, y) = if rotated {
                let scaled_x = left + (x - left) * scale;
                let dx = scaled_x - pivot_x;
                let dy = y - pivot_y;
                (pivot_x + dx * cos - dy * sin, pivot_y + dx * sin + dy * cos)
            } else {
                (left + (x - left) * scale, y)
            };
            if !x.is_finite() || !y.is_finite() {
                continue;
            }
            bounds[0] = bounds[0].min(x);
            bounds[1] = bounds[1].min(y);
            bounds[2] = bounds[2].max(x);
            bounds[3] = bounds[3].max(y);
        }
    }
    bounds
        .iter()
        .all(|value| value.is_finite())
        .then_some(bounds)
}

/// Clockwise rotation in degrees, or `None` when it is absent or a no-op.
fn rotation_degrees(rotation: Option<&serde_json::Number>) -> Option<f64> {
    let degrees = rotation.and_then(|value| value.as_f64())?;
    (degrees.is_finite() && degrees.abs() > f64::EPSILON).then_some(degrees)
}

/// Axis-aligned bounds of a box rotated around its centre, in CSS pixels.
fn rotated_box(rect: [f64; 4], rotation: Option<&serde_json::Number>) -> [f64; 4] {
    let Some(degrees) = rotation_degrees(rotation) else {
        return rect;
    };
    let center_x = (rect[0] + rect[2]) / 2.0;
    let center_y = (rect[1] + rect[3]) / 2.0;
    rotate_box(rect, degrees, (center_x, center_y))
}

/// Axis-aligned bounds of `rect` rotated clockwise about `center`.
fn rotate_box(rect: [f64; 4], degrees: f64, center: (f64, f64)) -> [f64; 4] {
    let (center_x, center_y) = center;
    let (sin, cos) = degrees.to_radians().sin_cos();
    let mut bounds = [
        f64::INFINITY,
        f64::INFINITY,
        f64::NEG_INFINITY,
        f64::NEG_INFINITY,
    ];
    for (x, y) in [
        (rect[0], rect[1]),
        (rect[2], rect[1]),
        (rect[2], rect[3]),
        (rect[0], rect[3]),
    ] {
        let dx = x - center_x;
        let dy = y - center_y;
        let rx = center_x + dx * cos - dy * sin;
        let ry = center_y + dx * sin + dy * cos;
        bounds[0] = bounds[0].min(rx);
        bounds[1] = bounds[1].min(ry);
        bounds[2] = bounds[2].max(rx);
        bounds[3] = bounds[3].max(ry);
    }
    bounds
}

/// `[x, y, x + w, y + h]` in CSS pixels.
fn frame(x: f64, y: f64, w: f64, h: f64) -> Option<[f64; 4]> {
    let rect = [x, y, x + w, y + h];
    rect.iter().all(|value| value.is_finite()).then_some(rect)
}

/// The painted frame of an image in CSS pixels.
///
/// A `contentFrame` overrides the primitive's own `x`, `y`, `w` and `h` field
/// by field, matching the native raster's `image_frame`; the primitive's
/// fields remain the fallback for any axis the frame omits.
fn image_frame(image: &ImagePrimitive) -> Option<[f64; 4]> {
    let content_frame = image.attrs.content_frame.as_deref();
    let field = |value: &Option<serde_json::Number>| -> Option<f64> { value.as_ref()?.as_f64() };
    let x = content_frame
        .and_then(|frame| field(&frame.x))
        .or_else(|| image.x.as_f64())?;
    let y = content_frame
        .and_then(|frame| field(&frame.y))
        .or_else(|| image.y.as_f64())?;
    let w = content_frame
        .and_then(|frame| field(&frame.w))
        .or_else(|| image.w.as_f64())?;
    let h = content_frame
        .and_then(|frame| field(&frame.h))
        .or_else(|| image.h.as_f64())?;
    frame(x, y, w, h)
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

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::display_list::{
        ContentFrame, ImagePrimitive, PlacedGlyph, ShapePathCommand, ShapePrimitive,
        ShapeTransformPrimitive,
    };
    use serde_json::Number;

    fn number(value: f64) -> Number {
        Number::from_f64(value).unwrap()
    }

    fn one_page(primitives: Vec<Primitive>) -> DisplayPage {
        DisplayPage {
            page_index: 0,
            width: Number::from(816),
            height: Number::from(1056),
            content_bounds: None,
            column_bounds: Vec::new(),
            section_id: None,
            section_index: None,
            section_page_index: None,
            section_page_number: None,
            page_label: None,
            primitives,
            watermark_primitive_count: None,
            background: None,
            page_borders: Vec::new(),
            header: None,
            footer: None,
            note_areas: Vec::new(),
            reservations: Vec::new(),
            table_gridlines: Vec::new(),
            unbuilt: false,
            position_span: None,
            hf_parts: None,
        }
    }

    #[test]
    fn rotated_image_link_box_covers_the_painted_frame() {
        let boxed = rotated_box(
            [100.0, 100.0, 300.0, 120.0],
            Some(&serde_json::Number::from(90)),
        );
        assert!((boxed[0] - 190.0).abs() < 1e-9);
        assert!((boxed[2] - 210.0).abs() < 1e-9);
        assert!((boxed[1] - 10.0).abs() < 1e-9);
        assert!((boxed[3] - 210.0).abs() < 1e-9);
    }

    #[test]
    fn rotated_image_link_annotation_covers_the_rotated_painted_frame() {
        let image = ImagePrimitive {
            rel_id: "rId7".to_owned(),
            x: number(100.0),
            y: number(100.0),
            w: number(200.0),
            h: number(20.0),
            rotation_deg: Some(Number::from(90)),
            opacity: None,
            filter: None,
            decorative: false,
            crop: None,
            alt_text: None,
            relationship_id: None,
            attrs: DocAttrs {
                href: Some("https://example.com/logo".to_owned()),
                ..DocAttrs::default()
            },
        };
        let links = collect_links(&one_page(vec![Primitive::Image(image)]));
        assert_eq!(links.len(), 1, "the rotated image must stay clickable");
        let rect = links[0].rect;
        // A 150x15pt frame at (75, 702) rotated 90° about its centre paints
        // over [142.5, 634.5, 157.5, 784.5]; an unrotated box would be
        // [75, 702, 225, 717] and miss the painted frame entirely.
        for (actual, expected) in rect.iter().zip([142.5_f32, 634.5, 157.5, 784.5]) {
            assert!(
                (actual - expected).abs() < 0.01,
                "annotation rect {rect:?} should cover the rotated frame"
            );
        }
        assert_eq!(
            links[0].action,
            LinkAction::Uri("https://example.com/logo".to_owned())
        );
    }

    #[test]
    fn image_link_annotation_uses_the_content_frame_before_rotation() {
        // The primitive's own fields describe the pre-layout box; the
        // content frame is where the native raster actually paints. The link
        // must cover that painted frame, and still follow it through rotation.
        let mut image = ImagePrimitive {
            rel_id: "rId7".to_owned(),
            x: number(100.0),
            y: number(100.0),
            w: number(200.0),
            h: number(20.0),
            rotation_deg: Some(Number::from(90)),
            opacity: None,
            filter: None,
            decorative: false,
            crop: None,
            alt_text: None,
            relationship_id: None,
            attrs: DocAttrs {
                href: Some("https://example.com/frame".to_owned()),
                ..DocAttrs::default()
            },
        };
        image.attrs.content_frame = Some(Box::new(ContentFrame {
            x: Some(number(40.0)),
            y: Some(number(60.0)),
            w: Some(number(100.0)),
            h: Some(number(40.0)),
        }));
        let links = collect_links(&one_page(vec![Primitive::Image(image.clone())]));
        assert_eq!(links.len(), 1);
        // The 100x40 frame centred at (90, 80) rotates to [70, 30, 110, 130]
        // CSS pixels, which is [52.5, 694.5, 82.5, 769.5] PDF points on the
        // 1056 px page. The primitive's own box would give [67.5, 627, 82.5,
        // 792], so a fallback to those fields cannot pass.
        assert_rect(links[0].rect, [52.5, 694.5, 82.5, 769.5]);

        // An axis the frame omits still falls back to the primitive, exactly
        // as the raster's field-by-field frame does.
        image.rotation_deg = None;
        image.attrs.content_frame = Some(Box::new(ContentFrame {
            x: Some(number(40.0)),
            y: None,
            w: None,
            h: Some(number(60.0)),
        }));
        let links = collect_links(&one_page(vec![Primitive::Image(image)]));
        assert_eq!(links.len(), 1);
        // x = 40, y = 100 (fallback), w = 200 (fallback), h = 60.
        assert_rect(links[0].rect, [30.0, 672.0, 180.0, 717.0]);
    }

    fn glyph_run(href: &str) -> GlyphRunPrimitive {
        GlyphRunPrimitive {
            font_id: 0,
            size: 16.0,
            color: "#000000".to_owned(),
            text: "Hi".to_owned(),
            paint_text: None,
            glyphs: vec![
                PlacedGlyph {
                    id: 1,
                    x: 72.0,
                    y: 96.0,
                    cluster: 0,
                    advance: 10.0,
                    logical_order: None,
                    bidi_level: None,
                },
                PlacedGlyph {
                    id: 2,
                    x: 90.0,
                    y: 96.0,
                    cluster: 1,
                    advance: 5.0,
                    logical_order: None,
                    bidi_level: None,
                },
            ],
            echo: None,
            paint_clip: None,
            word_spacing: None,
            rtl: None,
            opacity: None,
            rotation_deg: None,
            horizontal_scale: None,
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs: DocAttrs {
                href: Some(href.to_owned()),
                ..DocAttrs::default()
            },
        }
    }

    fn assert_rect(actual: [f32; 4], expected: [f32; 4]) {
        for (actual, expected) in actual.iter().zip(expected) {
            assert!(
                (actual - expected).abs() < 0.01,
                "rect {actual:?} != expected {expected:?}"
            );
        }
    }

    #[test]
    fn glyph_run_link_box_scales_about_the_run_left_edge() {
        let mut run = glyph_run("https://example.com/run");
        run.horizontal_scale = Some(Number::from(200));
        let links = collect_links(&one_page(vec![Primitive::GlyphRun(run)]));
        assert_eq!(links.len(), 1);
        // Glyphs at 72..82 and 90..95 scale to 72..92 and 108..118; the run
        // box is 1.05em tall (0.8em above, 0.25em below the baseline).
        assert_rect(links[0].rect, [54.0, 717.0, 88.5, 729.6]);
    }

    #[test]
    fn rotated_glyph_run_link_box_matches_the_painter_contract() {
        let mut run = glyph_run("https://example.com/rotated");
        run.rotation_deg = Some(Number::from(90));
        let links = collect_links(&one_page(vec![Primitive::GlyphRun(run)]));
        assert_eq!(links.len(), 1, "the rotated run must stay clickable");
        // The painter scales about the run's left edge and rotates the
        // scaled glyph boxes about the glyph-rect centre (x=(72+95)/2,
        // first baseline - 0.2em) clockwise in display space. All four
        // glyph-box corners must stay inside the link.
        assert_rect(links[0].rect, [57.225, 713.775, 69.825, 731.025]);
    }

    #[test]
    fn rotated_glyph_run_link_box_scales_offsets_once() {
        let mut run = glyph_run("https://example.com/rotated-scaled");
        run.rotation_deg = Some(Number::from(90));
        run.horizontal_scale = Some(Number::from(200));
        let links = collect_links(&one_page(vec![Primitive::GlyphRun(run)]));
        assert_eq!(links.len(), 1);
        assert_rect(links[0].rect, [57.225, 696.525, 69.825, 731.025]);
    }

    #[test]
    fn rotated_shape_link_box_covers_the_transformed_frame() {
        let shape = ShapePrimitive {
            x: Number::from(100),
            y: Number::from(100),
            w: Number::from(200),
            h: Number::from(20),
            geometry_path: vec![ShapePathCommand::Move {
                x: Number::from(100),
                y: Number::from(100),
            }],
            fill: Some("#000000".to_owned()),
            stroke: None,
            transform: Some(ShapeTransformPrimitive {
                rotation: Some(Number::from(90)),
                flip_h: false,
                flip_v: false,
            }),
            decorative: false,
            attrs: DocAttrs {
                href: Some("https://example.com/shape".to_owned()),
                ..DocAttrs::default()
            },
        };
        let links = collect_links(&one_page(vec![Primitive::Shape(shape)]));
        assert_eq!(links.len(), 1);
        assert_rect(links[0].rect, [142.5, 634.5, 157.5, 784.5]);
    }

    #[test]
    fn rotated_text_link_box_matches_the_painter_baseline_pivot() {
        let text = TextRunPrimitive {
            text: "Rotated".to_owned(),
            paint_text: None,
            x: Number::from(100),
            baseline_y: Number::from(200),
            width: Number::from(48),
            paint_clip: None,
            font: "16px Roboto".to_owned(),
            scripts: None,
            color: "#000000".to_owned(),
            letter_spacing: None,
            word_spacing: None,
            rtl: None,
            opacity: None,
            rotation_deg: Some(Number::from(90)),
            horizontal_scale: None,
            substitute_scale: None,
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs: DocAttrs {
                href: Some("https://example.com/rotated-text".to_owned()),
                ..DocAttrs::default()
            },
        };
        let links = collect_links(&one_page(vec![Primitive::Text(text)]));
        assert_eq!(links.len(), 1);
        // The painter rotates the 48x16.8px text box about
        // (x + width / 2, baseline - 0.2em), not about the box centre.
        assert_rect(links[0].rect, [87.6, 626.4, 100.2, 662.4]);
    }

    #[test]
    fn merge_is_transitive_when_a_fragment_bridges_two_boxes() {
        let action = LinkAction::Uri("https://example.com".to_owned());
        let merged = merge_adjacent_links(vec![
            LinkAnnotation {
                rect: [0.0, 0.0, 10.0, 10.0],
                action: action.clone(),
                title: None,
            },
            LinkAnnotation {
                rect: [100.0, 0.0, 110.0, 10.0],
                action: action.clone(),
                title: None,
            },
            LinkAnnotation {
                rect: [10.0, 0.0, 100.0, 10.0],
                action,
                title: None,
            },
        ]);
        assert_eq!(merged.len(), 1, "{merged:?}");
        assert_eq!(merged[0].rect, [0.0, 0.0, 110.0, 10.0]);
    }

    #[test]
    fn unsupported_action_is_rejected_in_pdfa_mode() {
        let mut pdf = Pdf::new();
        let result = write_annotation(
            &mut pdf,
            Ref::new(1),
            Ref::new(2),
            &LinkAnnotation {
                rect: [10.0, 20.0, 30.0, 40.0],
                action: LinkAction::Unsupported("JavaScript"),
                title: None,
            },
            crate::PdfConformance::PdfA2u,
        );
        assert!(
            matches!(result, Err(PdfError::Unsupported("JavaScript"))),
            "{result:?}"
        );
    }
}
