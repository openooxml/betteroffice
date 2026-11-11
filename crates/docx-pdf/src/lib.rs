//! Native PDF export for DOCX display lists.
//!
//! This crate deliberately does not depend on the VSDX writer. It consumes the
//! renderer-agnostic [`DisplayList`] produced by `docx-layout` and emits a
//! searchable PDF with embedded TrueType subsets. It covers page geometry,
//! positioned glyph runs, fallback text, rectangles, lines, decorations,
//! images (RGB plus soft masks), shape paths, headers/footers, note areas and
//! page borders. Links and tagged PDF are follow-up work.

use std::collections::BTreeMap;

use docx_layout::display_list::{
    DecorationPrimitive, DisplayList, DisplayPage, GlyphRunPrimitive, ImagePrimitive,
    LinePrimitive, PageBorderPrimitive, PageBorderZOrder, Primitive, RectPrimitive,
    ShapePathCommand, ShapePrimitive, TextRunPrimitive,
};
use ooxml_text::FontStore;
use pdf_writer::types::{CidFontType, FontFlags, SystemInfo};
use pdf_writer::{Content, Finish, Name, Pdf, Rect, Ref, Str};
use serde_json::Number;

use crate::font::{EmbeddedFont, FontError};

mod font;
mod image;

/// PDF export failed.
#[derive(Debug)]
pub enum PdfError {
    /// The display list contains content this slice cannot represent yet.
    Unsupported(&'static str),
    /// A numeric field was not finite or outside the PDF range.
    InvalidNumber(&'static str),
    /// No pages were supplied.
    NoPages,
    /// A font could not be loaded or subset.
    Font(FontError),
    /// An image could not be decoded.
    Image(String),
}

impl From<FontError> for PdfError {
    fn from(error: FontError) -> Self {
        Self::Font(error)
    }
}

impl std::fmt::Display for PdfError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Unsupported(what) => write!(f, "unsupported PDF content: {what}"),
            Self::InvalidNumber(what) => write!(f, "invalid PDF number: {what}"),
            Self::NoPages => f.write_str("display list has no pages"),
            Self::Font(error) => write!(f, "{error}"),
            Self::Image(error) => write!(f, "image error: {error}"),
        }
    }
}

impl std::error::Error for PdfError {}

/// CSS pixels per PDF point (96 dpi / 72 dpi).
const PX_TO_PT: f32 = 72.0 / 96.0;
const FONT_NAME: Name<'static> = Name(b"F0");

#[derive(Clone)]
struct ImageRefs {
    resource: String,
    id: Ref,
    smask: Option<Ref>,
}

#[derive(Clone, Copy)]
struct FontRefs {
    type0: Ref,
    cid: Ref,
    descriptor: Ref,
    file: Ref,
    to_unicode: Ref,
}

/// Export a display list as a plain, untagged PDF.
///
/// Text primitives use a Base-14 font. Glyph runs require
/// [`export_pdf_with_fonts`].
pub fn export_pdf(display_list: &DisplayList) -> Result<Vec<u8>, PdfError> {
    export_pdf_inner(display_list, None, &BTreeMap::new())
}

/// Export a display list with embedded, subset TrueType fonts.
///
/// `store` must be the same font registry that produced the display list's
/// `font_id`s.
pub fn export_pdf_with_fonts(
    display_list: &DisplayList,
    store: &FontStore,
) -> Result<Vec<u8>, PdfError> {
    export_pdf_inner(display_list, Some(store), &BTreeMap::new())
}

/// Export with embedded fonts and registered image bytes keyed by relationship
/// id (or a `data:` URL in the primitive).
pub fn export_pdf_with_fonts_and_images(
    display_list: &DisplayList,
    store: &FontStore,
    images: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<u8>, PdfError> {
    export_pdf_inner(display_list, Some(store), images)
}

fn export_pdf_inner(
    display_list: &DisplayList,
    store: Option<&FontStore>,
    images: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<u8>, PdfError> {
    if display_list.pages.is_empty() {
        return Err(PdfError::NoPages);
    }

    let mut pdf = Pdf::new();
    let catalog_id = Ref::new(1);
    let page_tree_id = Ref::new(2);
    let text_font_id = Ref::new(3);
    let has_text = display_list.pages.iter().any(|page| {
        page.primitives
            .iter()
            .chain(
                page.header
                    .iter()
                    .flat_map(|region| region.primitives.iter())
                    .chain(
                        page.footer
                            .iter()
                            .flat_map(|region| region.primitives.iter()),
                    )
                    .chain(
                        page.note_areas
                            .iter()
                            .flat_map(|region| region.primitives.iter())
                            .chain(
                                page.note_areas
                                    .iter()
                                    .flat_map(|region| region.separator_primitives.iter()),
                            ),
                    ),
            )
            .any(|primitive| matches!(primitive, Primitive::Text(_)))
    });
    pdf.catalog(catalog_id).pages(page_tree_id);
    if has_text {
        pdf.type1_font(text_font_id).base_font(Name(b"Helvetica"));
    }

    let font_uses = font::collect(display_list);
    if store.is_none() && !font_uses.is_empty() {
        return Err(PdfError::Unsupported("glyph runs need embedded fonts"));
    }

    let mut font_map = BTreeMap::new();
    let mut font_objects = Vec::new();
    let mut next_id = if has_text { 4 } else { 3 };
    if let Some(store) = store {
        for (index, (font_id, glyphs)) in font_uses.iter().enumerate() {
            if glyphs.is_empty() {
                continue;
            }
            let resource = format!("F{}", index + 1);
            let embedded = EmbeddedFont::new(resource, store, *font_id, glyphs)?;
            let refs = FontRefs {
                type0: Ref::new(next_id),
                cid: Ref::new(next_id + 1),
                descriptor: Ref::new(next_id + 2),
                file: Ref::new(next_id + 3),
                to_unicode: Ref::new(next_id + 4),
            };
            next_id += 5;
            font_objects.push((*font_id, refs));
            font_map.insert(*font_id, embedded);
        }
    }

    for (font_id, refs) in &font_objects {
        write_font(&mut pdf, &font_map[font_id], *refs);
    }

    let mut image_map: BTreeMap<String, ImageRefs> = BTreeMap::new();
    let mut image_objects = Vec::new();
    if store.is_some() {
        for page in &display_list.pages {
            for primitive in page.primitives.iter().chain(
                page.header
                    .iter()
                    .flat_map(|region| region.primitives.iter())
                    .chain(
                        page.footer
                            .iter()
                            .flat_map(|region| region.primitives.iter()),
                    )
                    .chain(
                        page.note_areas
                            .iter()
                            .flat_map(|region| region.primitives.iter())
                            .chain(
                                page.note_areas
                                    .iter()
                                    .flat_map(|region| region.separator_primitives.iter()),
                            ),
                    ),
            ) {
                let Primitive::Image(image) = primitive else {
                    continue;
                };
                if image_map.contains_key(&image.rel_id) {
                    continue;
                }
                let Some(bytes) = image_bytes(images, &image.rel_id) else {
                    return Err(PdfError::Image(format!(
                        "no bytes registered for relationship {}",
                        image.rel_id
                    )));
                };
                let decoded =
                    image::DecodedImage::from_bytes(bytes.as_ref()).map_err(PdfError::Image)?;
                let smask = decoded.alpha.as_ref().map(|_| {
                    let id = Ref::new(next_id + 1);
                    id
                });
                let refs = ImageRefs {
                    resource: format!("Im{}", image_map.len() + 1),
                    id: Ref::new(next_id),
                    smask,
                };
                next_id += if smask.is_some() { 2 } else { 1 };
                image_map.insert(image.rel_id.clone(), refs.clone());
                image_objects.push((refs, decoded));
            }
        }
    }

    for (refs, decoded) in &image_objects {
        write_image(&mut pdf, refs, decoded);
    }

    let mut page_ids = Vec::with_capacity(display_list.pages.len());
    for page in &display_list.pages {
        let page_id = Ref::new(next_id);
        let content_id = Ref::new(next_id + 1);
        next_id += 2;

        let width = points(&page.width, "page width")?;
        let height = points(&page.height, "page height")?;
        let content = page_content(page, width, height, &font_map, &image_map)?;

        pdf.stream(content_id, &content);
        let mut writer = pdf.page(page_id);
        writer
            .parent(page_tree_id)
            .media_box(Rect::new(0.0, 0.0, width, height));
        writer.contents(content_id);
        let mut resources = writer.resources();
        let mut fonts = resources.fonts();
        if has_text {
            fonts.pair(FONT_NAME, text_font_id);
        }
        for (font_id, refs) in &font_objects {
            let resource = font_map[font_id].resource.as_bytes();
            fonts.pair(Name(resource), refs.type0);
        }
        fonts.finish();
        if !image_map.is_empty() {
            let mut xobjects = resources.x_objects();
            for refs in image_map.values() {
                xobjects.pair(Name(refs.resource.as_bytes()), refs.id);
            }
            xobjects.finish();
        }
        resources.finish();
        writer.finish();
        page_ids.push(page_id);
    }

    pdf.pages(page_tree_id)
        .kids(page_ids.iter().copied())
        .count(page_ids.len() as i32);
    Ok(pdf.finish())
}

fn write_font(pdf: &mut Pdf, font: &EmbeddedFont, refs: FontRefs) {
    let base = Name(font.base_name.as_bytes());
    pdf.type0_font(refs.type0)
        .base_font(base)
        .encoding_predefined(Name(b"Identity-H"))
        .descendant_font(refs.cid)
        .to_unicode(refs.to_unicode);

    {
        let mut cid = pdf.cid_font(refs.cid);
        cid.subtype(CidFontType::Type2)
            .base_font(base)
            .system_info(SystemInfo {
                registry: Str(b"Adobe"),
                ordering: Str(b"Identity"),
                supplement: 0,
            })
            .font_descriptor(refs.descriptor)
            .default_width(1000.0)
            .cid_to_gid_map_predefined(Name(b"Identity"));
        cid.widths().consecutive(0, font.widths.iter().copied());
    }

    {
        let mut descriptor = pdf.font_descriptor(refs.descriptor);
        descriptor
            .name(base)
            .flags(FontFlags::from_bits_retain(font.flags))
            .bbox(Rect::new(
                font.bbox[0],
                font.bbox[1],
                font.bbox[2],
                font.bbox[3],
            ))
            .italic_angle(font.italic_angle)
            .ascent(font.ascent)
            .descent(font.descent)
            .cap_height(font.cap_height)
            .stem_v(80.0)
            .font_file2(refs.file);
    }

    pdf.stream(refs.file, &font.subset);
    pdf.cmap(refs.to_unicode, &font.tounicode);
}

fn page_content(
    page: &DisplayPage,
    width: f32,
    height: f32,
    fonts: &BTreeMap<u32, EmbeddedFont>,
    images: &BTreeMap<String, ImageRefs>,
) -> Result<Vec<u8>, PdfError> {
    let mut content = Content::new();
    if let Some(background) = page.background.as_deref() {
        if let Some(rgb) = parse_color(background) {
            content.set_fill_rgb(rgb[0], rgb[1], rgb[2]);
            content.rect(0.0, 0.0, width, height);
            content.fill_nonzero();
        }
    }

    for border in page
        .page_borders
        .iter()
        .filter(|border| border.z_order == Some(PageBorderZOrder::Back))
    {
        paint_page_border(&mut content, page, border)?;
    }

    paint_primitives(&mut content, page, &page.primitives, fonts, images)?;
    if let Some(header) = &page.header {
        paint_primitives(&mut content, page, &header.primitives, fonts, images)?;
    }
    if let Some(footer) = &page.footer {
        paint_primitives(&mut content, page, &footer.primitives, fonts, images)?;
    }
    for note in &page.note_areas {
        paint_primitives(
            &mut content,
            page,
            &note.separator_primitives,
            fonts,
            images,
        )?;
        paint_primitives(&mut content, page, &note.primitives, fonts, images)?;
    }
    for border in page
        .page_borders
        .iter()
        .filter(|border| border.z_order != Some(PageBorderZOrder::Back))
    {
        paint_page_border(&mut content, page, border)?;
    }

    Ok(content.finish().into_vec())
}

fn paint_primitives(
    content: &mut Content,
    page: &DisplayPage,
    primitives: &[Primitive],
    fonts: &BTreeMap<u32, EmbeddedFont>,
    images: &BTreeMap<String, ImageRefs>,
) -> Result<(), PdfError> {
    for primitive in primitives {
        match primitive {
            Primitive::Text(text) => paint_text(content, page, text)?,
            Primitive::Rect(rect) => paint_rect(content, page, rect)?,
            Primitive::Line(line) => paint_line(content, page, line)?,
            Primitive::GlyphRun(run) => paint_glyph_run(content, page, run, fonts)?,
            Primitive::Image(image) => paint_image(content, page, image, images)?,
            Primitive::Shape(shape) => paint_shape(content, page, shape)?,
            Primitive::Decoration(deco) => paint_decoration(content, page, deco)?,
        }
    }
    Ok(())
}

fn paint_page_border(
    content: &mut Content,
    page: &DisplayPage,
    border: &PageBorderPrimitive,
) -> Result<(), PdfError> {
    let x = points(&border.x, "border x")?;
    let y = points(&border.y, "border y")?;
    let w = points(&border.w, "border width")?;
    let h = points(&border.h, "border height")?;
    if w <= 0.0 || h <= 0.0 {
        return Ok(());
    }
    let page_height = points(&page.height, "page height")?;
    for (side, x1, y1, x2, y2) in [
        (
            border.top.as_ref(),
            x,
            page_height - y,
            x + w,
            page_height - y,
        ),
        (
            border.bottom.as_ref(),
            x,
            page_height - y - h,
            x + w,
            page_height - y - h,
        ),
        (
            border.left.as_ref(),
            x,
            page_height - y,
            x,
            page_height - y - h,
        ),
        (
            border.right.as_ref(),
            x + w,
            page_height - y,
            x + w,
            page_height - y - h,
        ),
    ] {
        let Some(side) = side else {
            continue;
        };
        let Some(color) = parse_color(&side.color) else {
            continue;
        };
        let width = points(&side.width, "border side width")?;
        if width <= 0.0 {
            continue;
        }
        content.set_stroke_rgb(color[0], color[1], color[2]);
        content.set_line_width(width);
        content.move_to(x1, y1);
        content.line_to(x2, y2);
        content.stroke();
    }
    Ok(())
}

fn paint_text(
    content: &mut Content,
    page: &DisplayPage,
    text: &TextRunPrimitive,
) -> Result<(), PdfError> {
    if text.hidden || text.text.is_empty() {
        return Ok(());
    }
    if !text.text.is_ascii() {
        return Err(PdfError::Unsupported(
            "non-ASCII text needs an embedded font",
        ));
    }
    let size = font_size_px(&text.font) * PX_TO_PT;
    let x = points(&text.x, "text x")?;
    let baseline = points(&text.baseline_y, "text baseline")?;
    let page_height = points(&page.height, "page height")?;
    let color = parse_color(&text.color).unwrap_or([0.0, 0.0, 0.0]);

    content.begin_text();
    content.set_font(FONT_NAME, size);
    content.set_fill_rgb(color[0], color[1], color[2]);
    content.set_text_matrix([1.0, 0.0, 0.0, 1.0, x, page_height - baseline]);
    content.show(Str(text.text.as_bytes()));
    content.end_text();
    Ok(())
}

fn image_bytes<'a>(
    images: &'a BTreeMap<String, Vec<u8>>,
    key: &str,
) -> Option<std::borrow::Cow<'a, [u8]>> {
    if let Some(bytes) = images.get(key) {
        return Some(std::borrow::Cow::Borrowed(bytes));
    }
    image::data_url_bytes(key).map(std::borrow::Cow::Owned)
}

fn write_image(pdf: &mut Pdf, refs: &ImageRefs, image: &image::DecodedImage) {
    {
        let mut xobject = pdf.image_xobject(refs.id, &image.rgb);
        xobject
            .width(image.width)
            .height(image.height)
            .bits_per_component(8);
        xobject.color_space().device_rgb();
        if let Some(smask) = refs.smask {
            xobject.s_mask(smask);
        }
    }
    if let (Some(smask), Some(alpha)) = (refs.smask, image.alpha.as_ref()) {
        let mut mask = pdf.image_xobject(smask, alpha);
        mask.width(image.width)
            .height(image.height)
            .bits_per_component(8);
        mask.color_space().device_gray();
    }
}

fn paint_image(
    content: &mut Content,
    page: &DisplayPage,
    image: &ImagePrimitive,
    images: &BTreeMap<String, ImageRefs>,
) -> Result<(), PdfError> {
    let Some(refs) = images.get(&image.rel_id) else {
        return Err(PdfError::Image(format!(
            "no decoded image for relationship {}",
            image.rel_id
        )));
    };
    let x = points(&image.x, "image x")?;
    let y = points(&image.y, "image y")?;
    let w = points(&image.w, "image width")?;
    let h = points(&image.h, "image height")?;
    if w <= 0.0 || h <= 0.0 {
        return Ok(());
    }
    let page_height = points(&page.height, "page height")?;
    let rotation = image
        .rotation_deg
        .as_ref()
        .and_then(Number::as_f64)
        .unwrap_or(0.0) as f32;
    content.save_state();
    if rotation.abs() > f32::EPSILON {
        let cx = x + w / 2.0;
        let cy = page_height - y - h / 2.0;
        let (sin, cos) = (-rotation).to_radians().sin_cos();
        let a = cos * w;
        let b = sin * w;
        let c = -sin * h;
        let d = cos * h;
        content.transform([a, b, c, d, cx - (a + c) / 2.0, cy - (b + d) / 2.0]);
    } else {
        content.transform([w, 0.0, 0.0, h, x, page_height - y - h]);
    }
    content.x_object(Name(refs.resource.as_bytes()));
    content.restore_state();
    Ok(())
}

fn paint_decoration(
    content: &mut Content,
    page: &DisplayPage,
    deco: &DecorationPrimitive,
) -> Result<(), PdfError> {
    let x = points(&deco.x, "decoration x")?;
    let y = points(&deco.y, "decoration y")?;
    let w = points(&deco.w, "decoration width")?;
    let h = points(&deco.h, "decoration height")?;
    if w <= 0.0 || h <= 0.0 {
        return Ok(());
    }
    let Some(color) = parse_color(&deco.color) else {
        return Ok(());
    };
    let page_height = points(&page.height, "page height")?;
    if deco.dashed || deco.dotted {
        let thickness = h.max(0.75);
        content.set_stroke_rgb(color[0], color[1], color[2]);
        content.set_line_width(thickness);
        let unit = if deco.dotted {
            thickness
        } else {
            thickness * 2.0
        };
        content.set_dash_pattern([unit, unit], 0.0);
        content.move_to(x, page_height - (y + h / 2.0));
        content.line_to(x + w, page_height - (y + h / 2.0));
        content.stroke();
    } else {
        content.set_fill_rgb(color[0], color[1], color[2]);
        content.rect(x, page_height - y - h, w, h);
        content.fill_nonzero();
    }
    Ok(())
}

fn paint_shape(
    content: &mut Content,
    page: &DisplayPage,
    shape: &ShapePrimitive,
) -> Result<(), PdfError> {
    if shape.geometry_path.is_empty() {
        return Ok(());
    }
    let page_height = points(&page.height, "page height")?;
    content.save_state();
    if let Some(transform) = &shape.transform {
        let rotation = transform
            .rotation
            .as_ref()
            .and_then(Number::as_f64)
            .unwrap_or(0.0) as f32;
        if rotation.abs() > f32::EPSILON || transform.flip_h || transform.flip_v {
            let cx = points(&shape.x, "shape x")? + points(&shape.w, "shape width")? / 2.0;
            let cy = page_height
                - points(&shape.y, "shape y")?
                - points(&shape.h, "shape height")? / 2.0;
            let sx = if transform.flip_h { -1.0 } else { 1.0 };
            let sy = if transform.flip_v { -1.0 } else { 1.0 };
            let (sin, cos) = (-rotation).to_radians().sin_cos();
            // translate(cx,cy) * rotate * scale * translate(-cx,-cy)
            let a = cos * sx;
            let b = sin * sx;
            let c = -sin * sy;
            let d = cos * sy;
            let e = cx - a * cx - c * cy;
            let f = cy - b * cx - d * cy;
            content.transform([a, b, c, d, e, f]);
        }
    }

    let mut current = (0.0_f32, 0.0_f32);
    for command in &shape.geometry_path {
        match command {
            ShapePathCommand::Move { x, y } => {
                let x = points(x, "shape x")?;
                let y = page_height - points(y, "shape y")?;
                content.move_to(x, y);
                current = (x, y);
            }
            ShapePathCommand::Line { x, y } => {
                let x = points(x, "shape x")?;
                let y = page_height - points(y, "shape y")?;
                content.line_to(x, y);
                current = (x, y);
            }
            ShapePathCommand::Quad { cpx, cpy, x, y } => {
                let cx = points(cpx, "shape control x")?;
                let cy = page_height - points(cpy, "shape control y")?;
                let x = points(x, "shape x")?;
                let y = page_height - points(y, "shape y")?;
                // pdf-writer has no quadratic operator; convert exactly to a
                // cubic: controls at 2/3 from each endpoint toward the quad
                // control point.
                let c1x = current.0 + 2.0 / 3.0 * (cx - current.0);
                let c1y = current.1 + 2.0 / 3.0 * (cy - current.1);
                let c2x = x + 2.0 / 3.0 * (cx - x);
                let c2y = y + 2.0 / 3.0 * (cy - y);
                content.cubic_to(c1x, c1y, c2x, c2y, x, y);
                current = (x, y);
            }
            ShapePathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                let x = points(x, "shape x")?;
                let y = page_height - points(y, "shape y")?;
                content.cubic_to(
                    points(cp1x, "shape control x")?,
                    page_height - points(cp1y, "shape control y")?,
                    points(cp2x, "shape control x")?,
                    page_height - points(cp2y, "shape control y")?,
                    x,
                    y,
                );
                current = (x, y);
            }
            ShapePathCommand::Close => {
                content.close_path();
            }
        }
    }

    let fill = shape.fill.as_deref().and_then(parse_color);
    let stroke = shape
        .stroke
        .as_ref()
        .and_then(|stroke| parse_color(&stroke.color).map(|color| (color, stroke)));
    if let Some(color) = fill {
        content.set_fill_rgb(color[0], color[1], color[2]);
    }
    if let Some((color, stroke)) = stroke {
        content.set_stroke_rgb(color[0], color[1], color[2]);
        content.set_line_width(points(&stroke.width, "stroke width")?);
    }
    match (fill.is_some(), stroke.is_some()) {
        (true, true) => {
            content.fill_nonzero_and_stroke();
        }
        (true, false) => {
            content.fill_nonzero();
        }
        (false, true) => {
            content.stroke();
        }
        (false, false) => {}
    }
    content.restore_state();
    Ok(())
}

fn paint_glyph_run(
    content: &mut Content,
    page: &DisplayPage,
    run: &GlyphRunPrimitive,
    fonts: &BTreeMap<u32, EmbeddedFont>,
) -> Result<(), PdfError> {
    if run.hidden || run.glyphs.is_empty() {
        return Ok(());
    }
    let Some(font) = fonts.get(&run.font_id) else {
        return Err(PdfError::Unsupported("glyph run font was not embedded"));
    };
    let rotation = run
        .rotation_deg
        .as_ref()
        .and_then(Number::as_f64)
        .unwrap_or(0.0);
    if rotation.abs() > f64::EPSILON {
        return Err(PdfError::Unsupported("rotated glyph runs"));
    }
    let size = (run.size * f64::from(PX_TO_PT)) as f32;
    let color = parse_color(&run.color).unwrap_or([0.0, 0.0, 0.0]);
    let page_height = points(&page.height, "page height")?;
    let horizontal_scale = run
        .horizontal_scale
        .as_ref()
        .and_then(Number::as_f64)
        .unwrap_or(1.0) as f32;

    content.begin_text();
    content.set_font(Name(font.resource.as_bytes()), size);
    content.set_fill_rgb(color[0], color[1], color[2]);
    if (horizontal_scale - 1.0).abs() > f32::EPSILON {
        content.set_horizontal_scaling(horizontal_scale * 100.0);
    }
    for glyph in &run.glyphs {
        let Some(glyph_id) = font.glyph(glyph.id) else {
            return Err(PdfError::Unsupported("glyph missing from subset"));
        };
        let x = f64_to_points(glyph.x, "glyph x")?;
        let y = f64_to_points(glyph.y, "glyph y")?;
        content.set_text_matrix([1.0, 0.0, 0.0, 1.0, x, page_height - y]);
        content.show(Str(&glyph_id.to_be_bytes()));
    }
    content.end_text();
    Ok(())
}

fn paint_rect(
    content: &mut Content,
    page: &DisplayPage,
    rect: &RectPrimitive,
) -> Result<(), PdfError> {
    let x = points(&rect.x, "rect x")?;
    let y = points(&rect.y, "rect y")?;
    let w = points(&rect.w, "rect width")?;
    let h = points(&rect.h, "rect height")?;
    if w <= 0.0 || h <= 0.0 {
        return Ok(());
    }
    let Some(color) = parse_color(&rect.fill) else {
        return Ok(());
    };
    let page_height = points(&page.height, "page height")?;
    content.set_fill_rgb(color[0], color[1], color[2]);
    content.rect(x, page_height - y - h, w, h);
    content.fill_nonzero();
    Ok(())
}

fn paint_line(
    content: &mut Content,
    page: &DisplayPage,
    line: &LinePrimitive,
) -> Result<(), PdfError> {
    let x1 = points(&line.x1, "line x1")?;
    let y1 = points(&line.y1, "line y1")?;
    let x2 = points(&line.x2, "line x2")?;
    let y2 = points(&line.y2, "line y2")?;
    let width = points(&line.stroke_width, "line width")?;
    if width <= 0.0 {
        return Ok(());
    }
    let Some(color) = parse_color(&line.color) else {
        return Ok(());
    };
    let page_height = points(&page.height, "page height")?;
    content.set_stroke_rgb(color[0], color[1], color[2]);
    content.set_line_width(width);
    content.move_to(x1, page_height - y1);
    content.line_to(x2, page_height - y2);
    content.stroke();
    Ok(())
}

fn points(value: &Number, what: &'static str) -> Result<f32, PdfError> {
    let value = value.as_f64().ok_or(PdfError::InvalidNumber(what))?;
    let value = (value * f64::from(PX_TO_PT)) as f32;
    if !value.is_finite() {
        return Err(PdfError::InvalidNumber(what));
    }
    Ok(value)
}

fn f64_to_points(value: f64, what: &'static str) -> Result<f32, PdfError> {
    let value = (value * f64::from(PX_TO_PT)) as f32;
    if !value.is_finite() {
        return Err(PdfError::InvalidNumber(what));
    }
    Ok(value)
}

/// Best-effort CSS font-size parser for the fallback `text` primitive.
fn font_size_px(css: &str) -> f32 {
    css.split_whitespace()
        .find_map(|part| part.strip_suffix("px"))
        .and_then(|value| value.parse::<f32>().ok())
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(12.0)
}

/// Parse `#rgb`, `#rrggbb`, `rgb(...)` or `rgba(...)`; alpha is ignored.
fn parse_color(value: &str) -> Option<[f32; 3]> {
    let value = value.trim();
    if let Some(hex) = value.strip_prefix('#') {
        return match hex.len() {
            3 => Some([
                hex_channel(&hex[0..1])?,
                hex_channel(&hex[1..2])?,
                hex_channel(&hex[2..3])?,
            ]),
            6 => Some([
                hex_channel(&hex[0..2])?,
                hex_channel(&hex[2..4])?,
                hex_channel(&hex[4..6])?,
            ]),
            8 => Some([
                hex_channel(&hex[0..2])?,
                hex_channel(&hex[2..4])?,
                hex_channel(&hex[4..6])?,
            ]),
            _ => None,
        };
    }
    let inner = value
        .strip_prefix("rgb(")
        .or_else(|| value.strip_prefix("rgba("))
        .and_then(|value| value.strip_suffix(')'))?;
    let mut parts = inner.split(',').map(str::trim);
    let r = channel(parts.next()?)?;
    let g = channel(parts.next()?)?;
    let b = channel(parts.next()?)?;
    Some([r, g, b])
}

fn hex_channel(value: &str) -> Option<f32> {
    let value = u8::from_str_radix(value, 16).ok()?;
    Some(f32::from(value) / 255.0)
}

fn channel(value: &str) -> Option<f32> {
    let value = value.parse::<f32>().ok()?;
    Some((value / 255.0).clamp(0.0, 1.0))
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::display_list::{DocAttrs, PlacedGlyph};

    fn text(x: f64, y: f64, value: &str) -> Primitive {
        Primitive::Text(TextRunPrimitive {
            text: value.to_owned(),
            paint_text: None,
            x: Number::from_f64(x).unwrap(),
            baseline_y: Number::from_f64(y).unwrap(),
            width: Number::from(100),
            paint_clip: None,
            font: "12px Helvetica".to_owned(),
            scripts: None,
            color: "#112233".to_owned(),
            letter_spacing: None,
            word_spacing: None,
            rtl: None,
            opacity: None,
            rotation_deg: None,
            horizontal_scale: None,
            substitute_scale: None,
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs: DocAttrs::default(),
        })
    }

    #[test]
    fn writes_images_with_soft_masks() {
        let mut png = std::io::Cursor::new(Vec::new());
        let source = ::image::RgbaImage::from_pixel(2, 2, ::image::Rgba([200, 20, 10, 128]));
        ::image::DynamicImage::ImageRgba8(source)
            .write_to(&mut png, ::image::ImageFormat::Png)
            .expect("encode png");
        let mut images = BTreeMap::new();
        images.insert("rId7".to_owned(), png.into_inner());

        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
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
                primitives: vec![Primitive::Image(ImagePrimitive {
                    rel_id: "rId7".to_owned(),
                    x: Number::from(72),
                    y: Number::from(72),
                    w: Number::from(48),
                    h: Number::from(48),
                    rotation_deg: None,
                    opacity: None,
                    filter: None,
                    decorative: false,
                    crop: None,
                    alt_text: None,
                    relationship_id: None,
                    attrs: DocAttrs::default(),
                })],
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
            }],
        };

        let mut store = FontStore::default();
        let bytes =
            export_pdf_with_fonts_and_images(&display_list, &mut store, &images).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert!(pdf.contains("/Subtype /Image"));
        assert!(pdf.contains("/SMask"));
        assert!(pdf.contains("/DeviceRGB"));
        assert!(pdf.contains("/DeviceGray"));
    }

    #[test]
    fn embeds_and_subsets_glyph_runs() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../packages/fonts/assets/Roboto-Bold.ttf");
        let data = std::fs::read(path).expect("font fixture");
        let mut store = FontStore::default();
        let font_id = store.register(data.clone()).expect("register font");

        let face = ttf_parser::Face::parse(&data, 0).expect("parse font");
        let upem = f64::from(face.units_per_em());
        let size = 16.0;
        let text = "Hello";
        let mut x = 72.0;
        let mut glyphs = Vec::new();
        for (cluster, ch) in text.char_indices() {
            let glyph = face.glyph_index(ch).expect("glyph").0;
            let advance = f64::from(face.glyph_hor_advance(ttf_parser::GlyphId(glyph)).unwrap())
                * size
                / upem;
            glyphs.push(PlacedGlyph {
                id: u32::from(glyph),
                x,
                y: 96.0,
                cluster: cluster as u32,
                advance,
                logical_order: None,
                bidi_level: None,
            });
            x += advance;
        }

        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
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
                primitives: vec![Primitive::GlyphRun(GlyphRunPrimitive {
                    font_id: font_id.to_u32(),
                    size,
                    color: "#112233".to_owned(),
                    text: text.to_owned(),
                    paint_text: None,
                    glyphs,
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
                    attrs: DocAttrs::default(),
                })],
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
            }],
        };

        let bytes = export_pdf_with_fonts(&display_list, &store).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert!(pdf.contains("/FontFile2"));
        assert!(pdf.contains("/Identity-H"));
        assert!(pdf.contains("/ToUnicode"));
        assert!(pdf.contains("/BaseFont /AAAAAA+BO"));
    }

    #[test]
    fn writes_a4_page_with_text() {
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
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
                primitives: vec![text(72.0, 96.0, "Hello")],
                watermark_primitive_count: None,
                background: Some("#ffffff".to_owned()),
                page_borders: Vec::new(),
                header: None,
                footer: None,
                note_areas: Vec::new(),
                reservations: Vec::new(),
                table_gridlines: Vec::new(),
                unbuilt: false,
                position_span: None,
                hf_parts: None,
            }],
        };
        let bytes = export_pdf(&display_list).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert!(pdf.starts_with("%PDF-1."));
        assert!(pdf.contains("/MediaBox [0 0 612 792]"));
        assert!(pdf.contains("Hello"));
    }

    #[test]
    fn rejects_non_ascii_without_fonts() {
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
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
                primitives: vec![text(72.0, 96.0, "你好")],
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
            }],
        };
        assert!(matches!(
            export_pdf(&display_list),
            Err(PdfError::Unsupported(_))
        ));
    }
}
