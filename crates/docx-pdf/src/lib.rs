//! Native PDF export for DOCX display lists.
//!
//! This crate deliberately does not depend on the VSDX writer. It consumes the
//! renderer-agnostic [`DisplayList`] produced by `docx-layout` and emits a
//! searchable PDF with embedded TrueType subsets. It covers page geometry,
//! positioned glyph runs, fallback text, rectangles, lines, decorations,
//! images (RGB plus soft masks), shape paths, headers/footers, note areas and
//! page borders. Links and tagged PDF are follow-up work.

use std::collections::BTreeMap;
use std::io::Write;

use docx_layout::display_list::{
    DecorationPrimitive, DisplayList, DisplayPage, LinePrimitive, PageBorderPrimitive,
    PageBorderZOrder, Primitive, RectPrimitive, ShapePathCommand, ShapePrimitive, TextRunPrimitive,
};
use flate2::Compression;
use flate2::write::ZlibEncoder;
use ooxml_text::FontStore;
use pdf_writer::types::{CidFontType, FontFlags, SystemInfo};
use pdf_writer::{Content, Filter, Finish, Name, Pdf, Rect, Ref, Str};
use serde_json::Number;

use crate::font::{EmbeddedFont, FontError};

mod font;
mod image;
mod link;
mod text;
mod units;

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
pub(crate) struct ImageRefs {
    pub(crate) resource: String,
    pub(crate) id: Ref,
    pub(crate) smask: Option<Ref>,
}

#[derive(Clone, Copy)]
enum ImagePart<'a> {
    Body,
    HeaderFooter(&'a str),
    Notes(Option<&'a str>),
}

fn scoped_image_key(part: ImagePart<'_>, rel_id: &str) -> String {
    match part {
        ImagePart::Body => format!("\u{1f}{rel_id}"),
        ImagePart::HeaderFooter(region_rel_id) => format!("hf:{region_rel_id}\u{1f}{rel_id}"),
        ImagePart::Notes(kind) => {
            let part = if kind == Some("endnote") {
                "endnotes"
            } else {
                "footnotes"
            };
            format!("{part}\u{1f}{rel_id}")
        }
    }
}

fn image_refs<'a>(
    images: &'a BTreeMap<String, ImageRefs>,
    part: ImagePart<'_>,
    rel_id: &str,
) -> Option<&'a ImageRefs> {
    images
        .get(&scoped_image_key(part, rel_id))
        .or_else(|| images.get(rel_id))
}

#[derive(Clone, Copy)]
struct FontRefs {
    type0: Ref,
    cid: Ref,
    descriptor: Ref,
    file: Ref,
    to_unicode: Ref,
    cid_to_gid: Ref,
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

/// Export with embedded fonts and registered image bytes.
///
/// Keys are the writer's scoped relationship keys (`"\u{1f}rId5"` for body,
/// `"hf:<part>\u{1f}rId5"` for header/footer, `"footnotes\u{1f}rId5"` or
/// `"endnotes\u{1f}rId5"` for notes). An unscoped `rel_id` and a `data:` URL
/// in the primitive are accepted as fallbacks.
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
    let has_builtin_text = store.is_none_or(|store| !text::store_has_usable_face(store));
    let has_text = has_builtin_text
        && display_list.pages.iter().any(|page| {
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

    let font_uses = match store {
        Some(store) => font::collect_with_text(display_list, store),
        None => font::collect(display_list),
    };
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
                cid_to_gid: Ref::new(next_id + 5),
            };
            next_id += 6;
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
            for primitive in &page.primitives {
                collect_image(
                    primitive,
                    ImagePart::Body,
                    images,
                    &mut image_map,
                    &mut image_objects,
                    &mut next_id,
                );
            }
            if let Some(header) = &page.header {
                for primitive in &header.primitives {
                    collect_image(
                        primitive,
                        ImagePart::HeaderFooter(&header.r_id),
                        images,
                        &mut image_map,
                        &mut image_objects,
                        &mut next_id,
                    );
                }
            }
            if let Some(footer) = &page.footer {
                for primitive in &footer.primitives {
                    collect_image(
                        primitive,
                        ImagePart::HeaderFooter(&footer.r_id),
                        images,
                        &mut image_map,
                        &mut image_objects,
                        &mut next_id,
                    );
                }
            }
            for note in &page.note_areas {
                for primitive in note
                    .separator_primitives
                    .iter()
                    .chain(note.primitives.iter())
                {
                    collect_image(
                        primitive,
                        ImagePart::Notes(note.kind.as_deref()),
                        images,
                        &mut image_map,
                        &mut image_objects,
                        &mut next_id,
                    );
                }
            }
        }
    }

    for (refs, decoded) in &image_objects {
        write_image(&mut pdf, refs, decoded);
    }

    let mut opacity_map: BTreeMap<u32, (String, Ref)> = BTreeMap::new();
    for page in &display_list.pages {
        for primitive in all_primitives(page) {
            let Some(key) = primitive_opacity(primitive) else {
                continue;
            };
            if opacity_map.contains_key(&key) {
                continue;
            }
            let name = format!("GS{}", opacity_map.len() + 1);
            let id = Ref::new(next_id);
            next_id += 1;
            let alpha = key as f32 / 1000.0;
            pdf.ext_graphics(id)
                .non_stroking_alpha(alpha)
                .stroking_alpha(alpha);
            opacity_map.insert(key, (name, id));
        }
    }

    let mut page_ids = Vec::with_capacity(display_list.pages.len());
    for page in &display_list.pages {
        let page_id = Ref::new(next_id);
        let content_id = Ref::new(next_id + 1);
        next_id += 2;

        let links = link::collect_links(page);
        let annotation_ids: Vec<Ref> = (0..links.len())
            .map(|index| Ref::new(next_id + index as i32))
            .collect();
        next_id += links.len() as i32;

        let width = points(&page.width, "page width")?;
        let height = points(&page.height, "page height")?;
        let content = deflate(&page_content(
            page,
            width,
            height,
            &font_map,
            &image_map,
            store,
            &opacity_map,
        )?);
        pdf.stream(content_id, &content).filter(Filter::FlateDecode);
        let mut writer = pdf.page(page_id);
        writer
            .parent(page_tree_id)
            .media_box(Rect::new(0.0, 0.0, width, height));
        writer.contents(content_id);
        if !annotation_ids.is_empty() {
            writer.annotations(annotation_ids.iter().copied());
        }
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
        if !opacity_map.is_empty() {
            let mut states = resources.ext_g_states();
            for (name, id) in opacity_map.values() {
                states.pair(Name(name.as_bytes()), *id);
            }
            states.finish();
        }
        resources.finish();
        writer.finish();
        for (id, annotation) in annotation_ids.iter().zip(&links) {
            link::write_annotation(&mut pdf, *id, page_id, annotation);
        }
        page_ids.push(page_id);
    }

    pdf.pages(page_tree_id)
        .kids(page_ids.iter().copied())
        .count(page_ids.len() as i32);
    Ok(pdf.finish())
}

fn deflate(data: &[u8]) -> Vec<u8> {
    let mut encoder = ZlibEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(data).expect("in-memory deflate");
    encoder.finish().expect("in-memory deflate")
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
        let subtype = if font.is_cff {
            CidFontType::Type0
        } else {
            CidFontType::Type2
        };
        cid.subtype(subtype)
            .base_font(base)
            .system_info(SystemInfo {
                registry: Str(b"Adobe"),
                ordering: Str(b"Identity"),
                supplement: 0,
            })
            .font_descriptor(refs.descriptor)
            .default_width(1000.0);
        if font.cid_to_gid.is_some() {
            cid.cid_to_gid_map_stream(refs.cid_to_gid);
        } else if !font.is_cff {
            cid.cid_to_gid_map_predefined(Name(b"Identity"));
        }
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
            .stem_v(80.0);
        if font.is_cff {
            descriptor.font_file3(refs.file);
        } else {
            descriptor.font_file2(refs.file);
        }
    }

    let subset = deflate(&font.subset);
    let mut stream = pdf.stream(refs.file, &subset);
    stream.filter(Filter::FlateDecode);
    if font.is_cff {
        stream.pair(Name(b"Subtype"), Name(b"OpenType"));
    } else {
        stream.pair(Name(b"Length1"), font.subset.len() as i32);
    }
    drop(stream);
    if let Some(map) = &font.cid_to_gid {
        let mut bytes = Vec::with_capacity(map.len() * 2);
        for gid in map {
            bytes.extend_from_slice(&gid.to_be_bytes());
        }
        let bytes = deflate(&bytes);
        let mut stream = pdf.stream(refs.cid_to_gid, &bytes);
        stream.filter(Filter::FlateDecode);
    }
    pdf.cmap(refs.to_unicode, &font.tounicode);
}

fn page_content(
    page: &DisplayPage,
    width: f32,
    height: f32,
    fonts: &BTreeMap<u32, EmbeddedFont>,
    images: &BTreeMap<String, ImageRefs>,
    store: Option<&FontStore>,
    opacities: &BTreeMap<u32, (String, Ref)>,
) -> Result<Vec<u8>, PdfError> {
    let mut content = Content::new();
    if let Some(rgb) = page.background.as_deref().and_then(parse_color) {
        content.set_fill_rgb(rgb[0], rgb[1], rgb[2]);
        content.rect(0.0, 0.0, width, height);
        content.fill_nonzero();
    }

    for border in page
        .page_borders
        .iter()
        .filter(|border| border.z_order == Some(PageBorderZOrder::Back))
    {
        paint_page_border(&mut content, page, border)?;
    }

    let watermark_count = page
        .watermark_primitive_count
        .unwrap_or(0)
        .min(page.primitives.len());
    let (watermark, body) = page.primitives.split_at(watermark_count);
    paint_primitives(
        &mut content,
        page,
        watermark,
        fonts,
        images,
        store,
        opacities,
        ImagePart::Body,
    )?;
    if let Some(header) = &page.header {
        paint_primitives(
            &mut content,
            page,
            &header.primitives,
            fonts,
            images,
            store,
            opacities,
            ImagePart::HeaderFooter(&header.r_id),
        )?;
    }
    if let Some(footer) = &page.footer {
        paint_primitives(
            &mut content,
            page,
            &footer.primitives,
            fonts,
            images,
            store,
            opacities,
            ImagePart::HeaderFooter(&footer.r_id),
        )?;
    }
    paint_primitives(
        &mut content,
        page,
        body,
        fonts,
        images,
        store,
        opacities,
        ImagePart::Body,
    )?;
    for note in &page.note_areas {
        paint_primitives(
            &mut content,
            page,
            &note.separator_primitives,
            fonts,
            images,
            store,
            opacities,
            ImagePart::Notes(note.kind.as_deref()),
        )?;
        paint_primitives(
            &mut content,
            page,
            &note.primitives,
            fonts,
            images,
            store,
            opacities,
            ImagePart::Notes(note.kind.as_deref()),
        )?;
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

fn all_primitives(page: &DisplayPage) -> impl Iterator<Item = &Primitive> {
    page.primitives.iter().chain(
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
}

fn primitive_opacity(primitive: &Primitive) -> Option<u32> {
    let value = match primitive {
        Primitive::Text(text) => text.opacity.as_ref(),
        Primitive::GlyphRun(run) => run.opacity.as_ref(),
        Primitive::Image(image) => {
            return image::opacity_alpha(image)
                .map(|alpha| (alpha * 1000.0).round() as u32)
                .filter(|key| *key < 1000);
        }
        _ => None,
    }?;
    let alpha = value.as_f64()?;
    if !alpha.is_finite() {
        return None;
    }
    let alpha = alpha.clamp(0.0, 1.0);
    if alpha >= 1.0 {
        return None;
    }
    Some((alpha * 1000.0).round() as u32)
}

fn paint_primitives(
    content: &mut Content,
    page: &DisplayPage,
    primitives: &[Primitive],
    fonts: &BTreeMap<u32, EmbeddedFont>,
    images: &BTreeMap<String, ImageRefs>,
    store: Option<&FontStore>,
    opacities: &BTreeMap<u32, (String, Ref)>,
    part: ImagePart<'_>,
) -> Result<(), PdfError> {
    for primitive in primitives {
        let opacity =
            primitive_opacity(primitive).and_then(|key| opacities.get(&key).map(|(name, _)| name));
        if let Some(name) = opacity {
            content.save_state();
            content.set_parameters(Name(name.as_bytes()));
        }
        match primitive {
            Primitive::Text(text) => match store {
                Some(store) => text::paint_text(content, page, text, fonts, store)?,
                None => paint_text(content, page, text)?,
            },
            Primitive::Rect(rect) => paint_rect(content, page, rect)?,
            Primitive::Line(line) => paint_line(content, page, line)?,
            Primitive::GlyphRun(run) => match store {
                Some(store) => text::paint_glyph_run_with_store(content, page, run, fonts, store)?,
                None => text::paint_glyph_run(content, page, run, fonts)?,
            },
            Primitive::Image(image) => {
                if let Some(refs) = image_refs(images, part, &image.rel_id) {
                    let page_height = points(&page.height, "page height")?;
                    image::paint_image(content, page_height, image, refs)?;
                }
            }
            Primitive::Shape(shape) => paint_shape(content, page, shape)?,
            Primitive::Decoration(deco) => paint_decoration(content, page, deco)?,
        }
        if opacity.is_some() {
            content.restore_state();
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

fn collect_image(
    primitive: &Primitive,
    part: ImagePart<'_>,
    images: &BTreeMap<String, Vec<u8>>,
    image_map: &mut BTreeMap<String, ImageRefs>,
    image_objects: &mut Vec<(ImageRefs, image::DecodedImage)>,
    next_id: &mut i32,
) {
    let Primitive::Image(image) = primitive else {
        return;
    };
    let key = scoped_image_key(part, &image.rel_id);
    if image_map.contains_key(&key) {
        return;
    }
    let Some(bytes) = image_bytes(images, part, &image.rel_id) else {
        return;
    };
    let Some(decoded) = image::try_decode(bytes.as_ref()) else {
        return;
    };
    let smask = decoded.alpha.as_ref().map(|_| Ref::new(*next_id + 1));
    let refs = ImageRefs {
        resource: format!("Im{}", image_map.len() + 1),
        id: Ref::new(*next_id),
        smask,
    };
    *next_id += if smask.is_some() { 2 } else { 1 };
    image_map.insert(key, refs.clone());
    image_objects.push((refs, decoded));
}

fn image_bytes<'a>(
    images: &'a BTreeMap<String, Vec<u8>>,
    part: ImagePart<'_>,
    rel_id: &str,
) -> Option<std::borrow::Cow<'a, [u8]>> {
    if let Some(bytes) = images.get(&scoped_image_key(part, rel_id)) {
        return Some(std::borrow::Cow::Borrowed(bytes));
    }
    if let Some(bytes) = images.get(rel_id) {
        return Some(std::borrow::Cow::Borrowed(bytes));
    }
    image::data_url_bytes(rel_id).map(std::borrow::Cow::Owned)
}

fn write_image(pdf: &mut Pdf, refs: &ImageRefs, image: &image::DecodedImage) {
    let rgb = deflate(&image.rgb);
    {
        let mut xobject = pdf.image_xobject(refs.id, &rgb);
        xobject
            .width(image.width)
            .height(image.height)
            .bits_per_component(8);
        xobject.color_space().device_rgb();
        xobject.filter(Filter::FlateDecode);
        if let Some(smask) = refs.smask {
            xobject.s_mask(smask);
        }
    }
    if let (Some(smask), Some(alpha)) = (refs.smask, image.alpha.as_ref()) {
        let alpha = deflate(alpha);
        let mut mask = pdf.image_xobject(smask, &alpha);
        mask.width(image.width)
            .height(image.height)
            .bits_per_component(8);
        mask.color_space().device_gray();
        mask.filter(Filter::FlateDecode);
    }
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
        content.save_state();
        content.set_dash_pattern([unit, unit], 0.0);
        content.move_to(x, page_height - (y + h / 2.0));
        content.line_to(x + w, page_height - (y + h / 2.0));
        content.stroke();
        content.restore_state();
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
            .unwrap_or(0.0);
        if !rotation.is_finite() {
            return Err(PdfError::InvalidNumber("shape rotation"));
        }
        let rotation = rotation as f32;
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
        (false, false) => {
            content.end_path();
        }
    }
    content.restore_state();
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
                hex_channel(hex.get(0..1)?)?,
                hex_channel(hex.get(1..2)?)?,
                hex_channel(hex.get(2..3)?)?,
            ]),
            6 => Some([
                hex_channel(hex.get(0..2)?)?,
                hex_channel(hex.get(2..4)?)?,
                hex_channel(hex.get(4..6)?)?,
            ]),
            8 => Some([
                hex_channel(hex.get(0..2)?)?,
                hex_channel(hex.get(2..4)?)?,
                hex_channel(hex.get(4..6)?)?,
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
    value.is_finite().then(|| (value / 255.0).clamp(0.0, 1.0))
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::display_list::{
        DecoKind, DocAttrs, GlyphRunPrimitive, HfKind, HfRegion, ImagePrimitive, PlacedGlyph,
    };

    fn number(value: f64) -> Number {
        Number::from_f64(value).unwrap()
    }

    /// One page carrying `primitives`, with `watermark_primitive_count` and a
    /// single header (footer stays empty).
    fn one_page(
        primitives: Vec<Primitive>,
        watermark_primitive_count: Option<usize>,
        header: Option<HfRegion>,
    ) -> DisplayList {
        DisplayList {
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
                primitives,
                watermark_primitive_count,
                background: None,
                page_borders: Vec::new(),
                header,
                footer: None,
                note_areas: Vec::new(),
                reservations: Vec::new(),
                table_gridlines: Vec::new(),
                unbuilt: false,
                position_span: None,
                hf_parts: None,
            }],
        }
    }

    fn rect(x: f64, y: f64, w: f64, h: f64, fill: &str) -> Primitive {
        Primitive::Rect(RectPrimitive {
            x: number(x),
            y: number(y),
            w: number(w),
            h: number(h),
            fill: fill.to_owned(),
            attrs: DocAttrs::default(),
        })
    }

    fn line(x1: f64, y1: f64, x2: f64, y2: f64, width: f64, color: &str) -> Primitive {
        Primitive::Line(LinePrimitive {
            x1: number(x1),
            y1: number(y1),
            x2: number(x2),
            y2: number(y2),
            stroke_width: number(width),
            color: color.to_owned(),
            dash: None,
            role: None,
            border_style: None,
            secondary_color: None,
            opacity: None,
            border_owner: None,
            attrs: DocAttrs::default(),
        })
    }

    /// A shape with nonempty geometry and neither fill nor stroke.
    fn paintless_shape(x: f64, y: f64) -> Primitive {
        Primitive::Shape(ShapePrimitive {
            x: number(x),
            y: number(y),
            w: number(100.0),
            h: number(50.0),
            geometry_path: vec![
                ShapePathCommand::Move {
                    x: number(10.0),
                    y: number(10.0),
                },
                ShapePathCommand::Line {
                    x: number(90.0),
                    y: number(10.0),
                },
                ShapePathCommand::Line {
                    x: number(90.0),
                    y: number(40.0),
                },
                ShapePathCommand::Close,
            ],
            fill: None,
            stroke: None,
            transform: None,
            decorative: false,
            attrs: DocAttrs::default(),
        })
    }

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
                    opacity: Some(Number::from_f64(0.5).unwrap()),
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
        assert!(pdf.contains("/ExtGState"));
        assert!(pdf.contains("/ca 0.5"));
    }

    #[test]
    fn zero_opacity_emits_a_transparent_graphics_state() {
        let mut page = text(72.0, 96.0, "transparent");
        if let Primitive::Text(run) = &mut page {
            run.opacity = Some(Number::from_f64(0.0).unwrap());
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
                primitives: vec![page],
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
        let bytes = export_pdf_with_fonts_and_images(&display_list, &mut store, &BTreeMap::new())
            .expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert!(pdf.contains("/ca 0"));
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

    fn inflate_first_stream(pdf: &[u8]) -> String {
        use std::io::Read;
        let start = pdf
            .windows(b"stream\n".len())
            .position(|window| window == b"stream\n")
            .expect("stream")
            + b"stream\n".len();
        let end = pdf[start..]
            .windows(b"endstream".len())
            .position(|window| window == b"endstream")
            .expect("endstream")
            + start;
        let mut decoded = String::new();
        flate2::read::ZlibDecoder::new(&pdf[start..end])
            .read_to_string(&mut decoded)
            .expect("inflate content");
        decoded
    }

    fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
        haystack
            .windows(needle.len())
            .position(|window| window == needle)
    }

    /// Decompress the content stream the page dictionary's `/Contents` names.
    ///
    /// Unlike [`inflate_first_stream`], this finds the page stream even when an
    /// image or font stream precedes it.
    fn inflate_page_content(pdf: &[u8]) -> String {
        use std::io::Read;
        let contents = find_bytes(pdf, b"/Contents ").expect("/Contents");
        let id_start = contents + b"/Contents ".len();
        let id_len = pdf[id_start..]
            .iter()
            .take_while(|byte| byte.is_ascii_digit())
            .count();
        assert!(id_len > 0, "content stream reference");
        let id = std::str::from_utf8(&pdf[id_start..id_start + id_len]).expect("content id");
        let object = format!("\n{id} 0 obj");
        let object_start = find_bytes(pdf, object.as_bytes()).expect("content object");
        let stream_start = object_start
            + find_bytes(&pdf[object_start..], b"stream\n").expect("content stream")
            + b"stream\n".len();
        let stream_end = stream_start
            + find_bytes(&pdf[stream_start..], b"endstream").expect("content stream end");
        let mut decoded = String::new();
        flate2::read::ZlibDecoder::new(&pdf[stream_start..stream_end])
            .read_to_string(&mut decoded)
            .expect("inflate content");
        decoded
    }

    /// Value of one `/ca` or `/CA` entry in the ExtGState dictionaries.
    fn ext_g_state_alpha(pdf: &str, key: &str) -> Option<f32> {
        let start = pdf.find(key)? + key.len();
        let rest = pdf[start..].trim_start();
        let end = rest
            .find(|character: char| {
                !(character.is_ascii_digit() || character == '.' || character == '-')
            })
            .unwrap_or(rest.len());
        rest[..end].parse().ok()
    }

    #[derive(Debug, Clone, Copy, Default)]
    struct GState {
        fill: Option<[f32; 3]>,
        stroke: Option<[f32; 3]>,
        dashed: bool,
    }

    enum Operand {
        Num(f32),
        Array(Vec<f32>),
    }

    /// One path-painting operation together with the graphics state and the
    /// path it actually painted.
    #[derive(Debug, Clone, PartialEq)]
    struct PaintOp {
        /// `n`, `S`/`s`, `f`/`F`/`f*`, `B`/`B*`/`b`/`b*`.
        op: String,
        fill: Option<[f32; 3]>,
        stroke: Option<[f32; 3]>,
        dashed: bool,
        /// Subpaths started since the previous paint; a leaked shape path
        /// shows up as an extra subpath on the next fill.
        subpaths: usize,
    }

    fn operands_rgb(operands: &[Operand]) -> Option<[f32; 3]> {
        match operands {
            [Operand::Num(r), Operand::Num(g), Operand::Num(b)] => Some([*r, *g, *b]),
            _ => None,
        }
    }

    /// Interpret the path, color, dash and `q`/`Q` operators of a page content
    /// stream and return its painting operations in order.
    ///
    /// Text and image operators are ignored: their operands never change the
    /// state this tracks (colors are RGB, dash arrays are the only arrays).
    fn paint_ops(content: &str) -> Vec<PaintOp> {
        let normalized = content.replace('[', " [ ").replace(']', " ] ");
        let mut tokens = normalized.split_whitespace();
        let mut operands: Vec<Operand> = Vec::new();
        let mut state = GState::default();
        let mut stack: Vec<GState> = Vec::new();
        let mut subpaths = 0usize;
        let mut ops = Vec::new();
        while let Some(token) = tokens.next() {
            if token == "[" {
                let mut values = Vec::new();
                for item in tokens.by_ref() {
                    if item == "]" {
                        break;
                    }
                    if let Ok(value) = item.parse::<f32>() {
                        values.push(value);
                    }
                }
                operands.push(Operand::Array(values));
                continue;
            }
            if let Ok(value) = token.parse::<f32>() {
                operands.push(Operand::Num(value));
                continue;
            }
            match token {
                "q" => stack.push(state),
                "Q" => state = stack.pop().unwrap_or_default(),
                "rg" => state.fill = operands_rgb(&operands),
                "RG" => state.stroke = operands_rgb(&operands),
                "d" => {
                    state.dashed = matches!(
                        operands.first(),
                        Some(Operand::Array(values)) if values.iter().any(|value| *value > 0.0)
                    );
                }
                "m" | "l" | "re" => subpaths += 1,
                "n" | "S" | "s" | "f" | "F" | "f*" | "B" | "B*" | "b" | "b*" => {
                    ops.push(PaintOp {
                        op: token.to_owned(),
                        fill: state.fill,
                        stroke: state.stroke,
                        dashed: state.dashed,
                        subpaths,
                    });
                    subpaths = 0;
                }
                _ => {}
            }
            operands.clear();
        }
        ops
    }

    #[test]
    fn paintless_shape_ends_its_path_before_the_next_fill() {
        let display_list = one_page(
            vec![
                paintless_shape(72.0, 120.0),
                rect(100.0, 120.0, 40.0, 30.0, "#ff0000"),
            ],
            None,
            None,
        );
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let first = ops.first().expect("a painting operation");
        assert_eq!(
            first.op, "n",
            "a paint-less shape must end its path before the next primitive"
        );
        let fill = ops
            .iter()
            .find(|op| op.op == "f")
            .expect("rect fill operation");
        assert_eq!(fill.fill, Some([1.0, 0.0, 0.0]));
        assert_eq!(
            fill.subpaths, 1,
            "the rect fill must paint only the rect, not the leaked shape path"
        );
    }

    #[test]
    fn watermark_paints_below_header_and_body_paints_above() {
        let header = HfRegion {
            r_id: "rIdHeader".to_owned(),
            kind: HfKind::Header,
            y: Number::from(0),
            height: Number::from(72),
            primitives: vec![rect(72.0, 36.0, 100.0, 10.0, "#00ff00")],
        };
        let display_list = one_page(
            vec![
                rect(0.0, 200.0, 300.0, 300.0, "#ff0000"),
                rect(72.0, 400.0, 100.0, 10.0, "#0000ff"),
            ],
            Some(1),
            Some(header),
        );
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let fills: Vec<Option<[f32; 3]>> = ops
            .iter()
            .filter(|op| op.op == "f")
            .map(|op| op.fill)
            .collect();
        assert_eq!(
            fills,
            vec![
                Some([1.0, 0.0, 0.0]),
                Some([0.0, 1.0, 0.0]),
                Some([0.0, 0.0, 1.0]),
            ],
            "watermark, header, then body must paint in that order"
        );
    }

    #[test]
    fn dashed_rule_does_not_leak_its_dash_into_a_solid_line() {
        let decoration = Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Underline,
            x: Number::from(72),
            y: Number::from(200),
            w: Number::from(100),
            h: Number::from(2),
            color: "#000000".to_owned(),
            dashed: true,
            dotted: false,
            attrs: DocAttrs::default(),
        });
        let display_list = one_page(
            vec![decoration, line(72.0, 240.0, 172.0, 240.0, 1.0, "#ff0000")],
            None,
            None,
        );
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let strokes: Vec<&PaintOp> = ops
            .iter()
            .filter(|op| op.op == "S" || op.op == "s")
            .collect();
        assert_eq!(strokes.len(), 2, "one dashed rule and one solid line");
        assert!(
            strokes[0].dashed,
            "the dashed decoration must set a dash pattern"
        );
        assert!(
            !strokes[1].dashed,
            "the solid line must not inherit the decoration dash pattern"
        );
        assert_eq!(strokes[1].stroke, Some([1.0, 0.0, 0.0]));
    }

    #[test]
    fn zero_opacity_image_emits_ca_zero() {
        let mut png = std::io::Cursor::new(Vec::new());
        let source = ::image::RgbaImage::from_pixel(2, 2, ::image::Rgba([10, 20, 30, 255]));
        ::image::DynamicImage::ImageRgba8(source)
            .write_to(&mut png, ::image::ImageFormat::Png)
            .expect("encode png");
        let mut images = BTreeMap::new();
        images.insert("rId9".to_owned(), png.into_inner());

        let display_list = one_page(
            vec![Primitive::Image(ImagePrimitive {
                rel_id: "rId9".to_owned(),
                x: Number::from(72),
                y: Number::from(72),
                w: Number::from(48),
                h: Number::from(48),
                rotation_deg: None,
                opacity: Some(number(0.0)),
                filter: None,
                decorative: false,
                crop: None,
                alt_text: None,
                relationship_id: None,
                attrs: DocAttrs::default(),
            })],
            None,
            None,
        );
        let mut store = FontStore::default();
        let bytes =
            export_pdf_with_fonts_and_images(&display_list, &mut store, &images).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert_eq!(ext_g_state_alpha(&pdf, "/ca "), Some(0.0));
        assert_eq!(ext_g_state_alpha(&pdf, "/CA "), Some(0.0));
        let content = inflate_page_content(&bytes);
        let state = content.find("/GS1 gs").expect("zero-alpha graphics state");
        let paint = content.find("/Im1 Do").expect("image paint");
        assert!(
            state < paint,
            "the image must paint inside the zero-alpha graphics state"
        );
    }

    #[test]
    fn rotated_image_link_annotation_covers_the_painted_frame() {
        let image = ImagePrimitive {
            rel_id: "rId7".to_owned(),
            x: Number::from(100),
            y: Number::from(100),
            w: Number::from(200),
            h: Number::from(20),
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
        let display_list = one_page(vec![Primitive::Image(image.clone())], None, None);
        let page = &display_list.pages[0];
        let links = crate::link::collect_links(page);
        assert_eq!(links.len(), 1, "the rotated image must stay clickable");
        let rect = links[0].rect;
        // Hand-computed: a 150x15pt frame at (75, 702) rotated 90° about its
        // centre paints over [142.5, 634.5, 157.5, 784.5]; an unrotated link
        // box would be [75, 702, 225, 717] and miss the painted frame.
        for (actual, expected) in rect.iter().zip([142.5_f32, 634.5, 157.5, 784.5]) {
            assert!(
                (actual - expected).abs() < 0.01,
                "annotation rect {rect:?} should cover the rotated frame"
            );
        }
        // Cross-check against the painter: every corner of the image unit
        // square, mapped through the placement the exporter uses, must lie
        // inside the clickable rectangle.
        let page_height = points(&page.height, "page height").expect("page height");
        let transform = crate::image::image_transform(&image, page_height).expect("transform");
        let [a, b, c, d, e, f] = transform.matrix;
        for (u, v) in [(0.0_f32, 0.0_f32), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)] {
            let x = a * u + c * v + e;
            let y = b * u + d * v + f;
            assert!(
                x >= rect[0] - 0.01
                    && x <= rect[2] + 0.01
                    && y >= rect[1] - 0.01
                    && y <= rect[3] + 0.01,
                "painted corner ({x}, {y}) must lie inside link rect {rect:?}"
            );
        }
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
        let content = inflate_first_stream(&bytes);
        assert!(content.contains("Hello"));
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
